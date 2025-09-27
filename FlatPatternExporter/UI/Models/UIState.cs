namespace FlatPatternExporter.UI.Models;

public class UIState
{
    private const string SCAN_BUTTON_TEXT = "Сканировать";
    private const string CANCEL_BUTTON_TEXT = "Прервать";
    private const string EXPORT_BUTTON_TEXT = "Экспорт";
    public const string CANCELLING_TEXT = "Прерывание...";

    public bool ScanEnabled { get; set; }
    public bool ExportEnabled { get; set; }
    public bool ClearEnabled { get; set; }
    public string ScanButtonText { get; set; } = "";
    public string ExportButtonText { get; set; } = "";
    public string ProgressText { get; set; } = "";
    public double ProgressValue { get; set; }
    public bool InventorUIDisabled { get; set; }
    public bool UpdateScanProgress { get; set; } = true;
    public bool UpdateExportProgress { get; set; } = true;

    public static UIState Initial => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "Документ не выбран",
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = true
    };

    public static UIState Scanning => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = CANCEL_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "Подготовка к сканированию...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = true,
        UpdateExportProgress = false
    };

    public static UIState CancellingScan => new()
    {
        ScanEnabled = false,
        ScanButtonText = CANCELLING_TEXT,
        ExportEnabled = false,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ClearEnabled = false,
        ProgressText = "Прерывание...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = true,
        UpdateExportProgress = false
    };

    public static UIState Exporting => new()
    {
        ScanEnabled = false,
        ExportEnabled = true,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = CANCEL_BUTTON_TEXT,
        ProgressText = "Экспорт данных...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = false,
        UpdateExportProgress = true
    };

    public static UIState CancellingExport => new()
    {
        ScanEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportEnabled = false,
        ExportButtonText = CANCELLING_TEXT,
        ClearEnabled = false,
        ProgressText = "Прерывание экспорта...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = false,
        UpdateExportProgress = true
    };

    public static UIState PreparingQuickExport => new()
    {
        ScanEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportEnabled = true,
        ExportButtonText = CANCEL_BUTTON_TEXT,
        ClearEnabled = false,
        ProgressText = "Сканирование и подготовка данных...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = false,
        UpdateExportProgress = true
    };

    public static UIState CreateClearedState() => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "",
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = false
    };

    public static UIState CreateAfterOperationState(bool hasData, bool wasCancelled, string statusText) => new()
    {
        ScanEnabled = true,
        ExportEnabled = hasData && !wasCancelled,
        ClearEnabled = hasData,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = statusText,
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = true
    };
}