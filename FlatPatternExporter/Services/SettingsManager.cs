using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.UI.Windows;

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
    public string SelectedLanguage { get; init; } = "en-US";
    public string SelectedTheme { get; init; } = "Dark";
}

public record ComponentFilterSettings
{
    public bool ExcludeReferenceParts { get; init; } = true;
    public bool ExcludePurchasedParts { get; init; } = true;
    public bool ExcludePhantomParts { get; init; } = true;
    public bool IncludeLibraryComponents { get; init; }
}

public record OrganizationSettings
{
    public bool OrganizeByMaterial { get; init; }
    public bool OrganizeByThickness { get; init; }
}

public record DxfExportSettings
{
    [JsonConverter(typeof(JsonStringEnumConverter))]
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
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public SplineReplacementType SelectedSplineReplacement { get; init; } = SplineReplacementType.Lines;
    public string SplineTolerance { get; init; } = DefaultSplineTolerance;
}

public record ExportFolderSettings
{
    [JsonConverter(typeof(JsonStringEnumConverter))]
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
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public CsvDelimiterType CsvDelimiter { get; init; } = CsvDelimiterType.Tab;

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ExportFileFormat DefaultExportFormat { get; init; } = ExportFileFormat.Excel;

    [JsonConverter(typeof(JsonStringEnumConverter))]
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

    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ProcessingMethod SelectedProcessingMethod { get; init; } = ProcessingMethod.BOM;

    public DxfExportSettings DxfExport { get; init; } = new();
    public SplineSettings Spline { get; init; } = new();
    public ExportFolderSettings ExportFolder { get; init; } = new();
    public FileNameSettings FileName { get; init; } = new();
    public ExcelExportSettings ExcelExport { get; init; } = new();
    public UpdateSettings Update { get; init; } = new();

    public List<LayerSettingData> LayerSettings { get; init; } = [];
}

public static class SettingsManager
{
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "FlatPatternExporter",
        "settings.json"
    );

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static ApplicationSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return new ApplicationSettings();
            }

            var json = File.ReadAllText(SettingsFilePath);
            return JsonSerializer.Deserialize<ApplicationSettings>(json, JsonOptions) ?? new ApplicationSettings();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error loading settings: {ex.Message}");
            return new ApplicationSettings();
        }
    }


    public static void SaveSettings(ApplicationSettings settings)
    {
        try
        {
            var directory = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error saving settings: {ex.Message}");
        }
    }

    public static ApplicationSettings CreateSettingsFromMainWindow(FlatPatternExporterMainWindow window)
    {
        var columnsInDisplayOrder = window.PartsDataGrid.Columns
            .Where(c => c.SortMemberPath is string { Length: > 0 })
            .OrderBy(c => c.DisplayIndex)
            .Select(c => c.SortMemberPath)
            .Where(internalName => !string.IsNullOrEmpty(internalName))
            .ToList();

        var templatePresets = window.PresetManager.GetPresetData().ToList();

        var layerSettings = window.LayerSettings
            .Where(ls => ls.HasChanges())
            .Select(layerSetting => new LayerSettingData
            {
                DisplayName = layerSetting.DisplayName,
                IsChecked = layerSetting.IsChecked,
                CustomName = layerSetting.CustomName,
                SelectedColor = layerSetting.SelectedColor,
                SelectedLineType = layerSetting.SelectedLineType
            })
            .ToList();

        var userDefinedProperties = PropertyMetadataRegistry.UserDefinedProperties
            .Select(p => p.InventorPropertyName ?? "")
            .Where(name => !string.IsNullOrEmpty(name))
            .ToList();

        var propertySubstitutions = window.PresetIProperties
            .Where(p => !string.IsNullOrEmpty(p.SubstitutionValue))
            .ToDictionary(p => p.InventorPropertyName, p => p.SubstitutionValue);

        return new ApplicationSettings
        {
            Interface = new InterfaceSettings
            {
                ColumnOrder = [..columnsInDisplayOrder],
                UserDefinedProperties = userDefinedProperties,
                PropertySubstitutions = propertySubstitutions,
                IsExpanded = window.SettingsExpander?.IsExpanded ?? false,
                SelectedLanguage = LocalizationManager.Instance.CurrentCulture.Name,
                SelectedTheme = window.ThemeToggleButton?.IsChecked == true ? "Dark" : "Light"
            },
            
            ComponentFilter = new ComponentFilterSettings
            {
                ExcludeReferenceParts = window.ExcludeReferenceParts,
                ExcludePurchasedParts = window.ExcludePurchasedParts,
                ExcludePhantomParts = window.ExcludePhantomParts,
                IncludeLibraryComponents = window.IncludeLibraryComponents
            },
            
            Organization = new OrganizationSettings
            {
                OrganizeByMaterial = window.OrganizeByMaterial,
                OrganizeByThickness = window.OrganizeByThickness
            },
            
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            
            DxfExport = new DxfExportSettings
            {
                SelectedAcadVersion = window.SelectedAcadVersion,
                MergeProfilesIntoPolyline = window.MergeProfilesIntoPolyline,
                RebaseGeometry = window.RebaseGeometry,
                TrimCenterlines = window.TrimCenterlines,
                OptimizeDxf = window.OptimizeDxf
            },
            
            Spline = new SplineSettings
            {
                EnableSplineReplacement = window.EnableSplineReplacement,
                SelectedSplineReplacement = window.SelectedSplineReplacement,
                SplineTolerance = window.SplineToleranceTextBox?.Text ?? SplineSettings.DefaultSplineTolerance
            },
            
            ExportFolder = new ExportFolderSettings
            {
                SelectedExportFolder = window.SelectedExportFolder,
                EnableSubfolder = window.EnableSubfolder,
                SubfolderName = window.SubfolderNameTextBox?.Text ?? "",
                FixedFolderPath = window.FixedFolderPath
            },
            
            FileName = new FileNameSettings
            {
                EnableFileNameConstructor = window.EnableFileNameConstructor,
                FileNameTemplate = window.TokenService.FileNameTemplate,
                TemplatePresets = templatePresets,
                SelectedTemplatePresetIndex = window.PresetManager.GetSelectedPresetIndex()
            },

            ExcelExport = new ExcelExportSettings
            {
                CsvDelimiter = window.CsvDelimiter,
                DefaultExportFormat = window.DefaultExportFormat,
                ExcelExportFileNameType = window.ExcelExportFileNameType
            },

            Update = new UpdateSettings
            {
                AutoUpdateCheck = window.AutoUpdateCheck
            },

            LayerSettings = layerSettings
        };
    }

    public static void ApplySettingsToMainWindow(ApplicationSettings settings, FlatPatternExporterMainWindow window)
    {
        // Interface settings
        if (window.SettingsExpander is not null)
            window.SettingsExpander.IsExpanded = settings.Interface.IsExpanded;

        // Restore language selection in ComboBox (language already applied in App.xaml.cs)
        var savedLanguage = SupportedLanguages.All
            .FirstOrDefault(lang => lang.Code == settings.Interface.SelectedLanguage);
        if (savedLanguage != null)
        {
            window.LanguageComboBox.SelectedItem = savedLanguage;
        }

        // Set ToggleButton state (theme already applied in App.xaml.cs)
        if (window.ThemeToggleButton is not null)
            window.ThemeToggleButton.IsChecked = settings.Interface.SelectedTheme == "Dark";

        // Component filter settings
        window.ExcludeReferenceParts = settings.ComponentFilter.ExcludeReferenceParts;
        window.ExcludePurchasedParts = settings.ComponentFilter.ExcludePurchasedParts;
        window.ExcludePhantomParts = settings.ComponentFilter.ExcludePhantomParts;
        window.IncludeLibraryComponents = settings.ComponentFilter.IncludeLibraryComponents;
        
        // Organization settings
        window.OrganizeByMaterial = settings.Organization.OrganizeByMaterial;
        window.OrganizeByThickness = settings.Organization.OrganizeByThickness;
        
        // Processing method
        window.SelectedProcessingMethod = settings.SelectedProcessingMethod;
        
        // DXF export settings
        window.SetAcadVersion(settings.DxfExport.SelectedAcadVersion, suppressMessage: true);
        window.MergeProfilesIntoPolyline = settings.DxfExport.MergeProfilesIntoPolyline;
        window.RebaseGeometry = settings.DxfExport.RebaseGeometry;
        window.TrimCenterlines = settings.DxfExport.TrimCenterlines;
        window.OptimizeDxf = settings.DxfExport.OptimizeDxf;
        
        // Spline settings
        window.EnableSplineReplacement = settings.Spline.EnableSplineReplacement;
        window.SelectedSplineReplacement = settings.Spline.SelectedSplineReplacement;
        if (window.SplineToleranceTextBox is not null)
            window.SplineToleranceTextBox.Text = settings.Spline.SplineTolerance;
        
        // Export folder settings
        window.SelectedExportFolder = settings.ExportFolder.SelectedExportFolder;
        window.EnableSubfolder = settings.ExportFolder.EnableSubfolder;
        if (window.SubfolderNameTextBox is not null)
            window.SubfolderNameTextBox.Text = settings.ExportFolder.SubfolderName;
        window.FixedFolderPath = settings.ExportFolder.FixedFolderPath;
        
        // File name settings
        window.EnableFileNameConstructor = settings.FileName.EnableFileNameConstructor;
        window.TokenService.FileNameTemplate = settings.FileName.FileNameTemplate;
        window.PresetManager.LoadPresets(settings.FileName.TemplatePresets, settings.FileName.SelectedTemplatePresetIndex);

        // Excel/CSV export settings
        window.CsvDelimiter = settings.ExcelExport.CsvDelimiter;
        window.DefaultExportFormat = settings.ExcelExport.DefaultExportFormat;
        window.ExcelExportFileNameType = settings.ExcelExport.ExcelExportFileNameType;

        // Update settings
        window.AutoUpdateCheck = settings.Update.AutoUpdateCheck;

        PropertyMetadataRegistry.UserDefinedProperties.Clear();

        // Load UDP from saved list
        foreach (var userProperty in settings.Interface.UserDefinedProperties)
        {
            if (!string.IsNullOrWhiteSpace(userProperty))
            {
                PropertyMetadataRegistry.AddUserDefinedProperty(userProperty);
            }
        }

        // Restore property substitutions directly to registry
        PropertyMetadataRegistry.PropertySubstitutions.Clear();
        foreach (var kvp in settings.Interface.PropertySubstitutions)
        {
            PropertyMetadataRegistry.PropertySubstitutions[kvp.Key] = kvp.Value;
        }

        // Apply substitutions to existing PresetIProperties
        foreach (var property in window.PresetIProperties)
        {
            if (PropertyMetadataRegistry.PropertySubstitutions.TryGetValue(property.InventorPropertyName, out var substitution))
            {
                property.SubstitutionValue = substitution;
            }
        }

        // Restore columns only if saved order exists
        if (settings.Interface.ColumnOrder.Count > 0)
        {
            var presetLookup = window.PresetIProperties.ToLookup(p => p.InventorPropertyName);

            var restoredColumns = 0;
            foreach (var internalName in settings.Interface.ColumnOrder.Where(name => !string.IsNullOrWhiteSpace(name)))
            {
                var propertyDef = PropertyMetadataRegistry.GetPropertyByInternalName(internalName);
                if (propertyDef == null)
                {
                    Debug.WriteLine($"Warning: Property '{internalName}' not found in registry");
                    continue;
                }

                if (PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
                {
                    window.AddUserDefinedIPropertyColumn(propertyDef.InventorPropertyName ?? internalName);
                    restoredColumns++;
                }
                else
                {
                    var presetProperty = presetLookup[internalName].FirstOrDefault();
                    if (presetProperty is not null)
                    {
                        window.AddIPropertyColumn(presetProperty);
                        restoredColumns++;
                    }
                    else
                    {
                        Debug.WriteLine($"Warning: Preset property '{internalName}' not found in PresetIProperties");
                    }
                }
            }

            Debug.WriteLine($"Restored {restoredColumns} columns out of {settings.Interface.ColumnOrder.Count} from settings");
        }
        else
        {
            Debug.WriteLine("Warning: No columns to restore from settings");
        }

        var layerSettingsLookup = window.LayerSettings.ToLookup(ls => ls.DisplayName);
        foreach (var settingData in settings.LayerSettings)
        {
            var layerSetting = layerSettingsLookup[settingData.DisplayName].FirstOrDefault();
            if (layerSetting is not null)
            {
                layerSetting.IsChecked = settingData.IsChecked;
                layerSetting.CustomName = settingData.CustomName;
                layerSetting.SelectedColor = settingData.SelectedColor;
                layerSetting.SelectedLineType = settingData.SelectedLineType;
            }
        }
    }

}
