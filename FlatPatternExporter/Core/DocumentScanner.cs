using System.Diagnostics;
using System.Runtime.InteropServices;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
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
        var result = new ScanResult();

        try
        {
            ClearCaches();
            _hasMissingReferences = false;

            var sheetMetalParts = new Dictionary<string, int>();

            if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                var asmDoc = (AssemblyDocument)document;

                if (processingMethod == ProcessingMethod.Traverse)
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, options, progress, cancellationToken), cancellationToken);
                else if (processingMethod == ProcessingMethod.BOM)
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, options, progress, cancellationToken), cancellationToken);

                await _conflictAnalyzer.AnalyzeConflictsAsync();
                _conflictAnalyzer.FilterConflictingParts(sheetMetalParts);

                result.SheetMetalParts = sheetMetalParts;
                result.ProcessedCount = sheetMetalParts.Count;
            }
            else if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                var partDoc = (PartDocument)document;

                if (!options.IncludeLibraryComponents && _inventorManager.IsLibraryComponent(partDoc.FullFileName))
                {
                    result.ProcessedCount = 0;
                    return result;
                }

                var mgr = new PropertyManager((Document)partDoc);
                var partNumber = mgr.GetMappedProperty("PartNumber");
                if (!string.IsNullOrEmpty(partNumber))
                {
                    _documentCache.AddDocumentToCache(partDoc, partNumber);

                    if (partDoc.SubType == PropertyManager.SheetMetalSubType)
                    {
                        sheetMetalParts.Add(partNumber, 1);
                        result.ProcessedCount = 1;
                        result.SheetMetalParts = sheetMetalParts;
                    }
                }
            }

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
        Dictionary<string, int> sheetMetalParts,
        ScanOptions options,
        IProgress<ScanProgress>? scanProgress = null,
        CancellationToken cancellationToken = default)
    {
        var filteredOccurrences = new List<ComponentOccurrence>();
        foreach (ComponentOccurrence occ in occurrences)
        {
            if (occ.Suppressed) continue;

            if (occ.Definition is VirtualComponentDefinition) continue;

            try
            {
                var fullFileName = GetFullFileName(occ);
                if (!ShouldExcludeComponent(occ.BOMStructure, fullFileName, options))
                    filteredOccurrences.Add(occ);
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
                    {
                        var mgr = new PropertyManager((Document)partDoc);
                        var partNumber = mgr.GetMappedProperty("PartNumber");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            _documentCache.AddDocumentToCache(partDoc, partNumber);

                            if (partDoc.SubType == PropertyManager.SheetMetalSubType)
                            {
                                var modelState = mgr.GetModelState();
                                _conflictAnalyzer.AddPartToTracker(partNumber, partDoc.FullFileName, modelState);

                                if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                    sheetMetalParts[partNumber]++;
                                else
                                    sheetMetalParts.Add(partNumber, 1);
                            }
                        }
                    }
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
                    CurrentOperation = "Сканирование компонентов",
                    CurrentItem = $"Компонент {processedOccurrences} из {totalOccurrences}"
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при обработке компонента: {ex.Message}");

                processedOccurrences++;
                scanProgress?.Report(new ScanProgress
                {
                    ProcessedItems = processedOccurrences,
                    TotalItems = totalOccurrences,
                    CurrentOperation = "Сканирование компонентов",
                    CurrentItem = $"Компонент {processedOccurrences} из {totalOccurrences}"
                });
            }
        }
    }

    private void ProcessBOM(
        BOM bom,
        Dictionary<string, int> sheetMetalParts,
        ScanOptions options,
        IProgress<ScanProgress>? scanProgress = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var allBomRows = GetAllBOMRowsRecursively(bom, options);

            Debug.WriteLine($"[ProcessBOM] Всего найдено {allBomRows.Count} строк BOM во всей структуре");

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
                        CurrentOperation = "Сканирование спецификации",
                        CurrentItem = $"Строка {processedRows} из {totalRows}"
                    });
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
                {
                    _hasMissingReferences = true;
                    Debug.WriteLine($"Обнаружен компонент с потерянной ссылкой: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка при обработке строки BOM: {ex.Message}");
                    _hasMissingReferences = true;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Общая ошибка при обработке BOM: {ex.Message}");
            _hasMissingReferences = true;
        }
    }

    private List<(BOMRow Row, int ParentQuantity)> GetAllBOMRowsRecursively(BOM bom, ScanOptions options, int parentQuantity = 1)
    {
        var allRows = new List<(BOMRow, int)>();

        try
        {
            var hideSuppressed = bom.HideSuppressedComponentsInBOM;

            BOMView? bomView = null;
            foreach (BOMView view in bom.BOMViews)
            {
                if (view.ViewType == BOMViewTypeEnum.kModelDataBOMViewType)
                {
                    bomView = view;
                    break;
                }
            }

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
                Debug.WriteLine("Ошибка при доступе к BOMRows. Возможно, в сборке есть потерянные ссылки.");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка при рекурсивном получении BOM: {ex.Message}");
        }

        return allRows;
    }

    private void ProcessBOMRowSimple(BOMRow row, Dictionary<string, int> sheetMetalParts, int parentQuantity = 1)
    {
        try
        {
            var componentDefinition = row.ComponentDefinitions[1];
            if (componentDefinition == null) return;

            if (componentDefinition.Document is not Document document) return;

            if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                if (document is PartDocument partDoc)
                {
                    var mgr = new PropertyManager((Document)partDoc);
                    var partNumber = mgr.GetMappedProperty("PartNumber");
                    if (!string.IsNullOrEmpty(partNumber))
                    {
                        _documentCache.AddDocumentToCache(partDoc, partNumber);

                        if (partDoc.SubType == PropertyManager.SheetMetalSubType)
                        {
                            var modelState = mgr.GetModelState();
                            _conflictAnalyzer.AddPartToTracker(partNumber, partDoc.FullFileName, modelState);

                            var totalQuantity = row.ItemQuantity * parentQuantity;

                            if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                sheetMetalParts[partNumber] += totalQuantity;
                            else
                                sheetMetalParts.Add(partNumber, totalQuantity);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка при простой обработке строки BOM: {ex.Message}");
        }
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

public class ScanResult
{
    public Dictionary<string, int> SheetMetalParts { get; set; } = [];
    public int ProcessedCount { get; set; }
    public int SkippedCount { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public bool WasCancelled { get; set; }
    public bool HasMissingReferences { get; set; }
    public List<string> Errors { get; set; } = [];
}

public class ScanOptions
{
    public bool ExcludeReferenceParts { get; set; }
    public bool ExcludePurchasedParts { get; set; }
    public bool ExcludePhantomParts { get; set; }
    public bool IncludeLibraryComponents { get; set; }
}