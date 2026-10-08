using System.Diagnostics;
using System.Runtime.InteropServices;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using Inventor;
using PropertyManager = FlatPatternExporter.Core.PropertyManager;

namespace FlatPatternExporter.Core;

public class DocumentScanner
{
    private readonly InventorManager _inventorManager;
    private readonly DocumentCache _documentCache;
    private readonly ConflictAnalyzer _conflictAnalyzer;
    private bool _hasMissingReferences;

    public DocumentScanner(InventorManager inventorManager)
    {
        _inventorManager = inventorManager;
        _documentCache = new DocumentCache();
        _conflictAnalyzer = new ConflictAnalyzer();
    }

    public DocumentCache DocumentCache => _documentCache;
    public ConflictAnalyzer ConflictAnalyzer => _conflictAnalyzer;
    public bool HasMissingReferences => _hasMissingReferences;

    public async Task<ScanResult> ScanDocumentAsync(
        Document document,
        ProcessingMethod processingMethod,
        ScanOptions options,
        IProgress<ScanProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new ScanResult
        {
            ProcessingMethod = processingMethod
        };

        try
        {
            ClearCaches();
            _hasMissingReferences = false;

            result.RootProperties = ReadRootProperties(document);

            var sheetMetalParts = new Dictionary<string, ScannedPart>(StringComparer.OrdinalIgnoreCase);

            if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                var asmDoc = (AssemblyDocument)document;

                if (processingMethod == ProcessingMethod.Traverse)
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, options, progress, cancellationToken), cancellationToken);
                else if (processingMethod == ProcessingMethod.BOM && options.BomView == BomViewType.PartsOnly)
                    await Task.Run(() => ProcessPartsOnlyBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, options, result.HiddenAssemblies, progress, cancellationToken), cancellationToken);
                else if (processingMethod == ProcessingMethod.BOM)
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, options, progress, cancellationToken), cancellationToken);

                await _conflictAnalyzer.AnalyzeConflictsAsync();
                if (!options.IncludeConflictingParts)
                    _conflictAnalyzer.FilterConflictingParts(sheetMetalParts);
            }
            else if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                var partDoc = (PartDocument)document;

                if (!options.IncludeLibraryComponents && _inventorManager.IsLibraryComponent(partDoc.FullFileName))
                    return result;

                ProcessPartDocument(partDoc, sheetMetalParts, 1);
            }

            result.SheetMetalParts = sheetMetalParts;
            result.WasCancelled = cancellationToken.IsCancellationRequested;
            result.HasMissingReferences = _hasMissingReferences;
        }
        catch (OperationCanceledException)
        {
            result.WasCancelled = true;
        }
        catch (Exception ex)
        {
            result.Errors.Add(ex.Message);
        }
        finally
        {
            stopwatch.Stop();
            result.ElapsedTime = stopwatch.Elapsed;
        }

        return result;
    }

    private void ProcessComponentOccurrences(
        ComponentOccurrences occurrences,
        Dictionary<string, ScannedPart> sheetMetalParts,
        ScanOptions options,
        IProgress<ScanProgress>? scanProgress = null,
        CancellationToken cancellationToken = default)
    {
        var filteredOccurrences = new List<ComponentOccurrence>();
        foreach (ComponentOccurrence occ in occurrences)
        {
            try
            {
                if (occ.Suppressed) continue;

                if (occ.Definition is VirtualComponentDefinition) continue;

                var fullFileName = GetFullFileName(occ);
                if (!ShouldExcludeComponent(occ.BOMStructure, fullFileName, options))
                    filteredOccurrences.Add(occ);
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                _hasMissingReferences = true;
                Debug.WriteLine($"Detected component with missing reference: {ex.Message}");
            }
            catch
            {
            }
        }

        var totalOccurrences = filteredOccurrences.Count;
        var processedOccurrences = 0;

        foreach (ComponentOccurrence occ in filteredOccurrences)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    if (occ.Definition.Document is PartDocument partDoc)
                        ProcessPartDocument(partDoc, sheetMetalParts, 1);
                }
                else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ProcessComponentOccurrences((ComponentOccurrences)occ.SubOccurrences, sheetMetalParts, options, cancellationToken: cancellationToken);
                }

                processedOccurrences++;
                scanProgress?.Report(new ScanProgress
                {
                    ProcessedItems = processedOccurrences,
                    TotalItems = totalOccurrences,
                    CurrentOperation = LocalizationManager.Instance.GetString("Status_ScanningComponents"),
                    CurrentItem = LocalizationManager.Instance.GetString("Status_ComponentProgress", processedOccurrences, totalOccurrences)
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing component: {ex.Message}");

                processedOccurrences++;
                scanProgress?.Report(new ScanProgress
                {
                    ProcessedItems = processedOccurrences,
                    TotalItems = totalOccurrences,
                    CurrentOperation = LocalizationManager.Instance.GetString("Status_ScanningComponents"),
                    CurrentItem = LocalizationManager.Instance.GetString("Status_ComponentProgress", processedOccurrences, totalOccurrences)
                });
            }
        }
    }

    private void ProcessBOM(
        BOM bom,
        Dictionary<string, ScannedPart> sheetMetalParts,
        ScanOptions options,
        IProgress<ScanProgress>? scanProgress = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var allBomRows = GetAllBOMRowsRecursively(bom, options);

            Debug.WriteLine($"[ProcessBOM] Total found {allBomRows.Count} BOM rows in entire structure");

            var processedRows = 0;
            var totalRows = allBomRows.Count;

            foreach (var (row, parentQuantity) in allBomRows)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    ProcessBOMRowSimple(row, sheetMetalParts, parentQuantity);

                    processedRows++;
                    scanProgress?.Report(new ScanProgress
                    {
                        ProcessedItems = processedRows,
                        TotalItems = totalRows,
                        CurrentOperation = LocalizationManager.Instance.GetString("Status_ScanningBOM"),
                        CurrentItem = LocalizationManager.Instance.GetString("Status_BomRowProgress", processedRows, totalRows)
                    });
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
                {
                    _hasMissingReferences = true;
                    Debug.WriteLine($"Detected component with missing reference: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing BOM row: {ex.Message}");
                    _hasMissingReferences = true;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"General error processing BOM: {ex.Message}");
            _hasMissingReferences = true;
        }
    }

    private List<(BOMRow Row, int ParentQuantity)> GetAllBOMRowsRecursively(BOM bom, ScanOptions options, int parentQuantity = 1)
    {
        var allRows = new List<(BOMRow, int)>();

        try
        {
            var hideSuppressed = bom.HideSuppressedComponentsInBOM;

            var bomView = GetBOMView(bom, BOMViewTypeEnum.kModelDataBOMViewType);
            if (bomView == null) return allRows;

            try
            {
                var bomRows = bomView.BOMRows.Cast<BOMRow>().ToArray();

                foreach (BOMRow row in bomRows)
                {
                    if (!hideSuppressed && row.ItemQuantity <= 0)
                        continue;

                    var componentDefinition = row.ComponentDefinitions[1];

                    if (componentDefinition is VirtualComponentDefinition)
                    {
                        continue;
                    }

                    if (componentDefinition?.Document is Document document &&
                        ShouldExcludeComponent(row.BOMStructure, document.FullFileName, options))
                        continue;

                    allRows.Add((row, parentQuantity));

                    try
                    {
                        if (componentDefinition?.Document is AssemblyDocument asmDoc)
                        {
                            var totalQuantity = parentQuantity * row.ItemQuantity;
                            var subRows = GetAllBOMRowsRecursively(asmDoc.ComponentDefinition.BOM, options, totalQuantity);
                            allRows.AddRange(subRows);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                _hasMissingReferences = true;
                Debug.WriteLine("Error accessing BOMRows. Assembly may have missing references.");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error recursively getting BOM: {ex.Message}");
        }

        return allRows;
    }

    private void ProcessPartsOnlyBOM(
        BOM bom,
        Dictionary<string, ScannedPart> sheetMetalParts,
        ScanOptions options,
        List<string> hiddenAssemblies,
        IProgress<ScanProgress>? scanProgress = null,
        CancellationToken cancellationToken = default)
    {
        var bomView = GetBOMView(bom, BOMViewTypeEnum.kPartsOnlyBOMViewType);
        if (bomView == null) return;

        var hideSuppressed = bom.HideSuppressedComponentsInBOM;
        var bomRows = bomView.BOMRows.Cast<BOMRow>().ToArray();
        var processedRows = 0;

        foreach (var row in bomRows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                ProcessPartsOnlyRow(row, hideSuppressed, sheetMetalParts, options, hiddenAssemblies);
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                _hasMissingReferences = true;
                Debug.WriteLine($"Detected component with missing reference: {ex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing parts only BOM row: {ex.Message}");
            }

            processedRows++;
            scanProgress?.Report(new ScanProgress
            {
                ProcessedItems = processedRows,
                TotalItems = bomRows.Length,
                CurrentOperation = LocalizationManager.Instance.GetString("Status_ScanningBOM"),
                CurrentItem = LocalizationManager.Instance.GetString("Status_BomRowProgress", processedRows, bomRows.Length)
            });
        }
    }

    private void ProcessPartsOnlyRow(
        BOMRow row,
        bool hideSuppressed,
        Dictionary<string, ScannedPart> sheetMetalParts,
        ScanOptions options,
        List<string> hiddenAssemblies)
    {
        var rowQuantity = row.ItemQuantity;
        if (!hideSuppressed && rowQuantity <= 0) return;

        var componentDefinition = row.ComponentDefinitions[1];
        if (componentDefinition is VirtualComponentDefinition) return;

        var bomStructure = row.BOMStructure;
        var document = componentDefinition.Document;

        // Inseparable and purchased assemblies are listed as single rows, their parts are not shown
        if (document is AssemblyDocument asmDoc)
        {
            if (!ShouldExcludeComponent(bomStructure, asmDoc.FullFileName, options) && ContainsSheetMetalParts(asmDoc))
                hiddenAssemblies.Add(new PropertyManager((Document)asmDoc).GetMappedProperty("PartNumber"));
            return;
        }

        List<(PartDocument PartDoc, int Quantity)> rowDocuments = row.Merged
            ? GetMergedRowDocuments(row)
            : document is PartDocument partDoc ? [(partDoc, rowQuantity)] : [];

        foreach (var (rowPartDoc, quantity) in rowDocuments)
        {
            if (!ShouldExcludeComponent(bomStructure, rowPartDoc.FullFileName, options))
                ProcessPartDocument(rowPartDoc, sheetMetalParts, quantity, row.ItemNumber);
        }
    }

    // A merged row combines different documents with the same part number: split it by referenced document
    private static List<(PartDocument PartDoc, int Quantity)> GetMergedRowDocuments(BOMRow row)
    {
        return [.. row.ComponentOccurrences.Cast<ComponentOccurrence>()
            .Select(occ => occ.Definition.Document)
            .OfType<PartDocument>()
            .GroupBy(partDoc => partDoc.FullDocumentName, StringComparer.OrdinalIgnoreCase)
            .Select(group => (group.First(), group.Count()))];
    }

    private static bool ContainsSheetMetalParts(AssemblyDocument asmDoc)
    {
        return asmDoc.AllReferencedDocuments.Cast<Document>()
            .Any(doc => doc.SubType == PropertyManager.SheetMetalSubType);
    }

    private static BOMView? GetBOMView(BOM bom, BOMViewTypeEnum viewType)
    {
        foreach (BOMView view in bom.BOMViews)
        {
            if (view.ViewType == viewType)
                return view;
        }
        return null;
    }

    public static bool IsPartsOnlyViewEnabled(Document document)
    {
        return document is not AssemblyDocument asmDoc || asmDoc.ComponentDefinition.BOM.PartsOnlyViewEnabled;
    }

    public static void EnablePartsOnlyView(Document document)
    {
        if (document is AssemblyDocument asmDoc)
            asmDoc.ComponentDefinition.BOM.PartsOnlyViewEnabled = true;
    }

    private void ProcessBOMRowSimple(BOMRow row, Dictionary<string, ScannedPart> sheetMetalParts, int parentQuantity = 1)
    {
        try
        {
            if (row.ComponentDefinitions[1]?.Document is PartDocument partDoc)
                ProcessPartDocument(partDoc, sheetMetalParts, row.ItemQuantity * parentQuantity);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error in simple BOM row processing: {ex.Message}");
        }
    }

    private void ProcessPartDocument(PartDocument partDoc, Dictionary<string, ScannedPart> sheetMetalParts, int quantity, string bomItem = "")
    {
        var key = partDoc.FullDocumentName;
        if (sheetMetalParts.TryGetValue(key, out var part))
        {
            part.Quantity += quantity;
            return;
        }

        // Already processed and not a sheet metal part
        if (_documentCache.GetCachedPartDocument(key) != null) return;

        var mgr = new PropertyManager((Document)partDoc);
        var partNumber = mgr.GetMappedProperty("PartNumber");
        if (string.IsNullOrEmpty(partNumber)) return;

        _documentCache.AddDocumentToCache(key, partDoc);

        if (partDoc.SubType != PropertyManager.SheetMetalSubType) return;

        _conflictAnalyzer.AddPartToTracker(partNumber, partDoc.FullFileName, mgr.GetModelState());
        sheetMetalParts.Add(key, new ScannedPart { FullDocumentName = key, PartNumber = partNumber, Quantity = quantity, BomItem = bomItem });
    }

    private static Dictionary<string, string> ReadRootProperties(Document document)
    {
        var mgr = new PropertyManager(document);
        return PropertyMetadataRegistry.RootProperties.Values
            .Select(p => p.SourceName)
            .ToDictionary(name => name, name => mgr.GetMappedProperty(name));
    }

    private bool ShouldExcludeComponent(BOMStructureEnum bomStructure, string fullFileName, ScanOptions options)
    {
        if (options.ExcludeReferenceParts && bomStructure == BOMStructureEnum.kReferenceBOMStructure)
            return true;

        if (options.ExcludePurchasedParts && bomStructure == BOMStructureEnum.kPurchasedBOMStructure)
            return true;

        if (options.ExcludePhantomParts && bomStructure == BOMStructureEnum.kPhantomBOMStructure)
            return true;

        if (!options.IncludeLibraryComponents && !string.IsNullOrEmpty(fullFileName) && _inventorManager.IsLibraryComponent(fullFileName))
            return true;

        return false;
    }

    private static string GetFullFileName(ComponentOccurrence occ)
    {
        string fullFileName = "";
        if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            if (occ.Definition.Document is PartDocument partDoc)
                fullFileName = partDoc.FullFileName;
        }
        else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            if (occ.Definition.Document is AssemblyDocument asmDoc)
                fullFileName = asmDoc.FullFileName;
        }
        return fullFileName;
    }

    public void ClearCaches()
    {
        _documentCache.ClearCache();
        _conflictAnalyzer.Clear();
    }
}

public class ScannedPart
{
    public string FullDocumentName { get; init; } = "";
    public string PartNumber { get; init; } = "";
    public int Quantity { get; set; }
    public string BomItem { get; init; } = "";
}

public class ScanResult
{
    public Dictionary<string, ScannedPart> SheetMetalParts { get; set; } = [];
    public IReadOnlyDictionary<string, string> RootProperties { get; set; } = new Dictionary<string, string>();
    public int ProcessedCount => SheetMetalParts.Count;
    public int SkippedCount { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public bool WasCancelled { get; set; }
    public bool HasMissingReferences { get; set; }
    public List<string> HiddenAssemblies { get; set; } = [];
    public ProcessingMethod ProcessingMethod { get; set; }
    public List<string> Errors { get; set; } = [];
}

public class ScanOptions
{
    public BomViewType BomView { get; set; }
    public bool ExcludeReferenceParts { get; set; }
    public bool ExcludePurchasedParts { get; set; }
    public bool ExcludePhantomParts { get; set; }
    public bool IncludeLibraryComponents { get; set; }
    public bool IncludeConflictingParts { get; set; }
}