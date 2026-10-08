using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Services;

public record LayerSettingData
{
    public string DisplayName { get; init; } = "";
    public bool IsChecked { get; init; }
    public string CustomName { get; init; } = LayerDefaults.DefaultCustomName;
    public string SelectedColor { get; init; } = LayerDefaults.DefaultColor;
    public string SelectedLineType { get; init; } = LayerDefaults.DefaultLineType;
}

public record TemplatePresetData
{
    public string Name { get; init; } = "";
    public string Template { get; init; } = "";
}

public record InterfaceSettings
{
    public List<string> ColumnOrder { get; init; } = [];
    public List<string> UserDefinedProperties { get; init; } = [];
    public Dictionary<string, string> PropertySubstitutions { get; init; } = [];
    public bool IsExpanded { get; init; }
    public string SelectedLanguage { get; init; } = SupportedLanguages.English.Code;
    public AppTheme SelectedTheme { get; init; } = AppTheme.Dark;
}

public record ComponentFilterSettings
{
    public bool ExcludeReferenceParts { get; init; } = true;
    public bool ExcludePurchasedParts { get; init; } = true;
    public bool ExcludePhantomParts { get; init; } = true;
    public bool IncludeLibraryComponents { get; init; }
    public bool IncludeConflictingParts { get; init; }
}

public record OrganizationSettings
{
    public bool OrganizeByMaterial { get; init; }
    public bool OrganizeByThickness { get; init; }
}

public record DxfExportSettings
{
    public AcadVersionType SelectedAcadVersion { get; init; } = AcadVersionType.V2000;
    public bool MergeProfilesIntoPolyline { get; init; } = true;
    public bool RebaseGeometry { get; init; } = true;
    public bool TrimCenterlines { get; init; }
    public bool OptimizeDxf { get; init; }
}

public record SplineSettings
{
    public const string DefaultSplineTolerance = "0.01";

    public bool EnableSplineReplacement { get; init; }
    public SplineReplacementType SelectedSplineReplacement { get; init; } = SplineReplacementType.Lines;
    public string SplineTolerance { get; init; } = DefaultSplineTolerance;
}

public record ExportFolderSettings
{
    public ExportFolderType SelectedExportFolder { get; init; } = ExportFolderType.ChooseFolder;
    public bool EnableSubfolder { get; init; }
    public string SubfolderName { get; init; } = "";
    public string FixedFolderPath { get; init; } = "";
}

public record FileNameSettings
{
    public bool EnableFileNameConstructor { get; init; }
    public string FileNameTemplate { get; init; } = "{PartNumber}";
    public List<TemplatePresetData> TemplatePresets { get; init; } = [];
    public int SelectedTemplatePresetIndex { get; init; } = -1;
}

public record ExcelExportSettings
{
    public CsvDelimiterType CsvDelimiter { get; init; } = CsvDelimiterType.Tab;
    public ExportFileFormat DefaultExportFormat { get; init; } = ExportFileFormat.Excel;
    public ExcelExportFileNameType ExcelExportFileNameType { get; init; } = ExcelExportFileNameType.DateTimeFormat;
}

public record UpdateSettings
{
    public bool AutoUpdateCheck { get; init; } = true;
}

public record ApplicationSettings
{
    public InterfaceSettings Interface { get; init; } = new();
    public ComponentFilterSettings ComponentFilter { get; init; } = new();
    public OrganizationSettings Organization { get; init; } = new();

    public ProcessingMethod SelectedProcessingMethod { get; init; } = ProcessingMethod.BOM;

    public DxfExportSettings DxfExport { get; init; } = new();
    public SplineSettings Spline { get; init; } = new();
    public ExportFolderSettings ExportFolder { get; init; } = new();
    public FileNameSettings FileName { get; init; } = new();
    public ExcelExportSettings ExcelExport { get; init; } = new();
    public UpdateSettings Update { get; init; } = new();

    public List<LayerSettingData> LayerSettings { get; init; } = [];
}
