using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;
using FlatPatternExporter.Utilities;
using Inventor;
using Path = System.IO.Path;
using File = System.IO.File;
using Directory = System.IO.Directory;

namespace FlatPatternExporter.Core;

public class ExportService
{
    private readonly InventorService _inventorService;
    private readonly ScanService _scanService;
    private readonly DocumentCacheService _documentCache;
    private readonly TokenService _tokenService;

    public ExportService(InventorService inventorService, ScanService scanService,
        DocumentCacheService documentCache, TokenService tokenService)
    {
        _inventorService = inventorService;
        _scanService = scanService;
        _documentCache = documentCache;
        _tokenService = tokenService;
    }

    public async Task<ExportContext> PrepareExportContextAsync(
        Document document,
        bool requireScan,
        bool showProgress,
        Document? lastScannedDocument,
        ExportOptions exportOptions)
    {
        var context = new ExportContext();

        try
        {
            if (requireScan && document != lastScannedDocument)
            {
                _scanService.ConflictAnalyzer.Clear();
            }

            if (requireScan && document != lastScannedDocument)
            {
                context.IsValid = false;
                context.ErrorMessage = "Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.";
                return context;
            }

            if (!PrepareForExport(exportOptions, out var targetDir, out var multiplier))
            {
                context.IsValid = false;
                context.ErrorMessage = "Ошибка при подготовке параметров экспорта";
                return context;
            }

            context.TargetDirectory = targetDir;
            context.Multiplier = multiplier;

            var scanOptions = new ScanOptions
            {
                ExcludeReferenceParts = exportOptions.ExcludeReferenceParts,
                ExcludePurchasedParts = exportOptions.ExcludePurchasedParts,
                ExcludePhantomParts = exportOptions.ExcludePhantomParts,
                IncludeLibraryComponents = exportOptions.IncludeLibraryComponents
            };

            var scanResult = await _scanService.ScanDocumentAsync(
                document,
                exportOptions.SelectedProcessingMethod,
                scanOptions,
                showProgress ? new Progress<ScanProgress>() : null);

            context.SheetMetalParts = scanResult.SheetMetalParts;
            context.IsValid = true;
        }
        catch (Exception ex)
        {
            context.IsValid = false;
            context.ErrorMessage = $"Ошибка при подготовке экспорта: {ex.Message}";
        }

        return context;
    }

    private bool PrepareForExport(ExportOptions exportOptions, out string targetDir, out int multiplier)
    {
        targetDir = "";
        multiplier = 1;

        switch (exportOptions.SelectedExportFolder)
        {
            case ExportFolderType.ChooseFolder:
                var dialog = new FolderBrowserDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                    targetDir = dialog.SelectedPath;
                else
                    return false;
                break;

            case ExportFolderType.ComponentFolder:
                targetDir = Path.GetDirectoryName(_inventorService.Application?.ActiveDocument?.FullFileName) ?? "";
                break;

            case ExportFolderType.FixedFolder:
                if (string.IsNullOrEmpty(exportOptions.FixedFolderPath))
                {
                    return false;
                }
                targetDir = exportOptions.FixedFolderPath;
                break;

            case ExportFolderType.ProjectFolder:
                try
                {
                    targetDir = _inventorService.Application?.DesignProjectManager?.ActiveDesignProject?.WorkspacePath ?? "";
                    if (string.IsNullOrEmpty(targetDir))
                    {
                        return false;
                    }
                }
                catch
                {
                    return false;
                }
                break;

            case ExportFolderType.PartFolder:
                break;
        }

        if (exportOptions.EnableSubfolder && !string.IsNullOrEmpty(exportOptions.SubfolderName))
        {
            targetDir = Path.Combine(targetDir!, exportOptions.SubfolderName);
            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
        }

        multiplier = exportOptions.Multiplier;
        return true;
    }

    public void ExportDXF(
        IEnumerable<PartData> partsDataList,
        string targetDir,
        int multiplier,
        ExportOptions exportOptions,
        ref int processedCount,
        ref int skippedCount,
        bool generateThumbnails,
        IProgress<double>? progress,
        CancellationToken cancellationToken = default)
    {
        var totalParts = partsDataList.Count();
        progress?.Report(0);

        var localProcessedCount = processedCount;
        var localSkippedCount = skippedCount;

        foreach (var partData in partsDataList)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var partNumber = partData.PartNumber;
            var qty = partData.IsOverridden ? partData.Quantity : partData.OriginalQuantity * multiplier;

            if (exportOptions.SelectedExportFolder == ExportFolderType.PartFolder)
            {
                var partPath = _documentCache.GetCachedPartPath(partNumber) ?? _inventorService.GetPartDocumentFullPath(partNumber);
                if (!string.IsNullOrEmpty(partPath))
                {
                    targetDir = Path.GetDirectoryName(partPath) ?? "";
                }
                else
                {
                    targetDir = "";
                }
            }

            PartDocument? partDoc = null;
            try
            {
                partDoc = _documentCache.GetCachedPartDocument(partNumber) ?? _inventorService.OpenPartDocument(partNumber);
                if (partDoc == null) throw new Exception("Файл детали не найден или не может быть открыт");

                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var material = partData.Material;
                var thickness = partData.Thickness;

                var materialDir = exportOptions.OrganizeByMaterial ? Path.Combine(targetDir, material) : targetDir;
                if (!Directory.Exists(materialDir)) Directory.CreateDirectory(materialDir);

                var thicknessDir = exportOptions.OrganizeByThickness
                    ? Path.Combine(materialDir, thickness)
                    : materialDir;
                if (!Directory.Exists(thicknessDir)) Directory.CreateDirectory(thicknessDir);

                string fileName;
                if (exportOptions.EnableFileNameConstructor && !string.IsNullOrEmpty(_tokenService.FileNameTemplate))
                {
                    fileName = _tokenService.ResolveTemplate(_tokenService.FileNameTemplate, partData);
                }
                else
                {
                    fileName = partNumber;
                }

                var filePath = Path.Combine(thicknessDir, fileName + ".dxf");

                if (!IsValidPath(filePath)) continue;

                var exportSuccess = false;

                if (smCompDef.HasFlatPattern)
                {
                    var flatPattern = smCompDef.FlatPattern;
                    var oDataIO = flatPattern.DataIO;

                    try
                    {
                        var options = PrepareExportOptions(exportOptions);

                        var layerOptionsBuilder = new StringBuilder();
                        var invisibleLayersBuilder = new StringBuilder();

                        foreach (var layer in exportOptions.LayerSettings)
                        {
                            if (layer.CanBeHidden && !layer.IsChecked)
                            {
                                invisibleLayersBuilder.Append($"{layer.LayerName};");
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(layer.CustomName))
                                layerOptionsBuilder.Append($"&{layer.DisplayName}={layer.CustomName}");

                            if (layer.SelectedLineType != "Default")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}LineType={LayerSettingsHelper.GetLineTypeValue(layer.SelectedLineType)}");

                            if (layer.SelectedColor != "White")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}Color={LayerSettingsHelper.GetColorValue(layer.SelectedColor)}");
                        }

                        if (invisibleLayersBuilder.Length > 0)
                        {
                            invisibleLayersBuilder.Length -= 1;
                            layerOptionsBuilder.Append($"&InvisibleLayers={invisibleLayersBuilder}");
                        }

                        options += layerOptionsBuilder.ToString();

                        var dxfOptions = $"FLAT PATTERN DXF?{options}";

                        Debug.WriteLine($"DXF Options: {dxfOptions}");

                        if (!File.Exists(filePath)) File.Create(filePath).Close();

                        if (IsFileLocked(filePath, exportOptions, cancellationToken))
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        oDataIO.WriteDataToFile(dxfOptions, filePath);
                        exportSuccess = true;

                        if (exportOptions.OptimizeDxf && exportSuccess)
                        {
                            if (AcadVersionMapping.SupportsOptimization(exportOptions.SelectedAcadVersion))
                            {
                                DxfOptimizer.OptimizeDxfFile(filePath, exportOptions.SelectedAcadVersion);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Ошибка экспорта DXF: {ex.Message}");
                    }
                }

                partData.ProcessingStatus = exportSuccess ? ProcessingStatus.Success : ProcessingStatus.Error;

                if (exportSuccess)
                    localProcessedCount++;
                else
                    localSkippedCount++;

                progress?.Report(totalParts > 0 ? (double)localProcessedCount / totalParts * 100 : 0);
            }
            catch (Exception ex)
            {
                localSkippedCount++;
                Debug.WriteLine($"Ошибка обработки детали: {ex.Message}");
            }
        }

        processedCount = localProcessedCount;
        skippedCount = localSkippedCount;

        progress?.Report(100);
    }

    private string PrepareExportOptions(ExportOptions exportOptions)
    {
        var sb = new StringBuilder();

        sb.Append($"AcadVersion={AcadVersionMapping.GetTranslatorCode(exportOptions.SelectedAcadVersion)}");

        if (exportOptions.EnableSplineReplacement)
        {
            var splineTolerance = exportOptions.SplineTolerance;

            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
            splineTolerance = splineTolerance.Replace('.', decimalSeparator).Replace(',', decimalSeparator);

            if (exportOptions.SelectedSplineReplacement == SplineReplacementType.Lines)
                sb.Append($"&SimplifySplines=True&SplineTolerance={splineTolerance}");
            else if (exportOptions.SelectedSplineReplacement == SplineReplacementType.Arcs)
                sb.Append($"&SimplifySplines=True&SimplifyAsTangentArcs=True&SplineTolerance={splineTolerance}");
        }
        else
        {
            sb.Append("&SimplifySplines=False");
        }

        if (exportOptions.MergeProfilesIntoPolyline) sb.Append("&MergeProfilesIntoPolyline=True");

        if (exportOptions.RebaseGeometry) sb.Append("&RebaseGeometry=True");

        if (exportOptions.TrimCenterlines)
            sb.Append("&TrimCenterlinesAtContour=True");

        return sb.ToString();
    }

    private bool IsFileLocked(string filePath, ExportOptions exportOptions, CancellationToken cancellationToken)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            stream.Close();
            return false;
        }
        catch (IOException)
        {
            if (exportOptions.ShowFileLockedDialogs)
            {
                var result = MessageBox.Show(
                    $"Файл {filePath} занят другим приложением. Прервать операцию?",
                    "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    throw new OperationCanceledException();
                }

                while (IsFileLockedInternal(filePath))
                {
                    var waitResult = MessageBox.Show(
                        "Ожидание разблокировки файла. Возможно файл открыт в программе просмотра. Закройте файл и нажмите OK.",
                        "Информация",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button2);

                    if (waitResult == DialogResult.Cancel)
                    {
                        throw new OperationCanceledException();
                    }

                    Thread.Sleep(1000);
                }

                cancellationToken.ThrowIfCancellationRequested();
            }
            return true;
        }
    }

    private static bool IsFileLockedInternal(string filePath)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            stream.Close();
            return false;
        }
        catch (IOException)
        {
            return true;
        }
    }

    private static bool IsValidPath(string path)
    {
        try
        {
            _ = new FileInfo(path);
            return true;
        }
        catch
        {
            return false;
        }
    }
}

public class ExportContext
{
    public string TargetDirectory { get; set; } = "";
    public int Multiplier { get; set; } = 1;
    public Dictionary<string, int> SheetMetalParts { get; set; } = [];
    public bool GenerateThumbnails { get; set; } = true;
    public bool IsValid { get; set; } = true;
    public string ErrorMessage { get; set; } = "";
}

public class ExportOptions
{
    public ExportFolderType SelectedExportFolder { get; set; }
    public string FixedFolderPath { get; set; } = "";
    public bool EnableSubfolder { get; set; }
    public string SubfolderName { get; set; } = "";
    public int Multiplier { get; set; } = 1;
    public ProcessingMethod SelectedProcessingMethod { get; set; }

    public bool ExcludeReferenceParts { get; set; }
    public bool ExcludePurchasedParts { get; set; }
    public bool ExcludePhantomParts { get; set; }
    public bool IncludeLibraryComponents { get; set; }

    public bool OrganizeByMaterial { get; set; }
    public bool OrganizeByThickness { get; set; }
    public bool EnableFileNameConstructor { get; set; }

    public bool OptimizeDxf { get; set; }
    public bool EnableSplineReplacement { get; set; }
    public SplineReplacementType SelectedSplineReplacement { get; set; }
    public string SplineTolerance { get; set; } = "0.01";
    public AcadVersionType SelectedAcadVersion { get; set; }
    public bool MergeProfilesIntoPolyline { get; set; }
    public bool RebaseGeometry { get; set; }
    public bool TrimCenterlines { get; set; }

    public List<LayerSetting> LayerSettings { get; set; } = [];
    public bool ShowFileLockedDialogs { get; set; } = true;
}
