using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FlatPatternExporter;

public record LayerSettingData
{
    public string DisplayName { get; init; } = "";
    public bool IsChecked { get; init; }
    public string CustomName { get; init; } = "";
    public string SelectedColor { get; init; } = "White";
    public string SelectedLineType { get; init; } = "Default";
}

public record TemplatePresetData
{
    public string Name { get; init; } = "";
    public string Template { get; init; } = "";
}

public record ApplicationSettings
{
    public const string DefaultSplineTolerance = "0.01";
    
    public List<string> ColumnOrder { get; init; } = [];
    public List<string> CustomProperties { get; init; } = [];
    
    public bool ExcludeReferenceParts { get; init; } = true;
    public bool ExcludePurchasedParts { get; init; } = true;
    public bool ExcludePhantomParts { get; init; } = true;
    public bool IncludeLibraryComponents { get; init; }
    
    public bool OrganizeByMaterial { get; init; }
    public bool OrganizeByThickness { get; init; }
    
    public bool EnableSplineReplacement { get; init; }
    public int SelectedSplineReplacementIndex { get; init; } = 0;
    public string SplineTolerance { get; init; } = DefaultSplineTolerance;
    public int SelectedAcadVersionIndex { get; init; } = 5;
    public bool MergeProfilesIntoPolyline { get; init; } = true;
    public bool RebaseGeometry { get; init; } = true;
    public bool TrimCenterlines { get; init; }
    
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ExportFolderType SelectedExportFolder { get; init; } = ExportFolderType.ChooseFolder;
    public bool EnableSubfolder { get; init; }
    public string SubfolderName { get; init; } = "";
    public string FixedFolderPath { get; init; } = "";
    
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ProcessingMethod SelectedProcessingMethod { get; init; } = ProcessingMethod.BOM;
    
    public bool OptimizeDxf { get; init; }
    
    public bool EnableFileNameConstructor { get; init; }
    public string FileNameTemplate { get; init; } = "{PartNumber}";
    
    public List<TemplatePresetData> TemplatePresets { get; init; } = [];
    public string SelectedTemplatePresetName { get; init; } = "";
    
    public List<LayerSettingData> LayerSettings { get; init; } = [];
    
    public bool IsExpanded { get; init; }
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
            .Where(c => c.Header is string headerValue && !string.IsNullOrEmpty(headerValue))
            .OrderBy(c => c.DisplayIndex)
            .Select(c => (string)c.Header)
            .ToList();

        var templatePresets = window.TemplatePresets
            .Select(preset => new TemplatePresetData
            {
                Name = preset.Name,
                Template = preset.Template
            })
            .ToList();

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

        return new ApplicationSettings
        {
            ColumnOrder = [..columnsInDisplayOrder],
            CustomProperties = [..window.CustomPropertiesList],
            
            ExcludeReferenceParts = window.ExcludeReferenceParts,
            ExcludePurchasedParts = window.ExcludePurchasedParts,
            ExcludePhantomParts = window.ExcludePhantomParts,
            IncludeLibraryComponents = window.IncludeLibraryComponents,
            
            OrganizeByMaterial = window.OrganizeByMaterial,
            OrganizeByThickness = window.OrganizeByThickness,
            
            EnableSplineReplacement = window.EnableSplineReplacement,
            SelectedSplineReplacementIndex = window.SelectedSplineReplacementIndex,
            SplineTolerance = window.SplineToleranceTextBox?.Text ?? ApplicationSettings.DefaultSplineTolerance,
            SelectedAcadVersionIndex = window.SelectedAcadVersionIndex,
            MergeProfilesIntoPolyline = window.MergeProfilesIntoPolyline,
            RebaseGeometry = window.RebaseGeometry,
            TrimCenterlines = window.TrimCenterlines,
            
            SelectedExportFolder = window.SelectedExportFolder,
            EnableSubfolder = window.EnableSubfolder,
            SubfolderName = window.SubfolderNameTextBox?.Text ?? "",
            FixedFolderPath = window.FixedFolderPath,
            
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            OptimizeDxf = window.OptimizeDxf,
            
            EnableFileNameConstructor = window.EnableFileNameConstructor,
            FileNameTemplate = window.TokenService.FileNameTemplate,
            
            TemplatePresets = templatePresets,
            SelectedTemplatePresetName = window.SelectedTemplatePreset?.Name ?? "",
            
            LayerSettings = layerSettings,
            IsExpanded = window.SettingsExpander?.IsExpanded ?? false
        };
    }

    public static void ApplySettingsToMainWindow(ApplicationSettings settings, FlatPatternExporterMainWindow window)
    {
        window.ExcludeReferenceParts = settings.ExcludeReferenceParts;
        window.ExcludePurchasedParts = settings.ExcludePurchasedParts;
        window.ExcludePhantomParts = settings.ExcludePhantomParts;
        window.IncludeLibraryComponents = settings.IncludeLibraryComponents;
        
        window.OrganizeByMaterial = settings.OrganizeByMaterial;
        window.OrganizeByThickness = settings.OrganizeByThickness;
        
        window.EnableSplineReplacement = settings.EnableSplineReplacement;
        window.SelectedSplineReplacementIndex = settings.SelectedSplineReplacementIndex;
        if (window.SplineToleranceTextBox is not null)
            window.SplineToleranceTextBox.Text = settings.SplineTolerance;
        window.SelectedAcadVersionIndex = settings.SelectedAcadVersionIndex;
        window.MergeProfilesIntoPolyline = settings.MergeProfilesIntoPolyline;
        window.RebaseGeometry = settings.RebaseGeometry;
        window.TrimCenterlines = settings.TrimCenterlines;
        
        window.SelectedExportFolder = settings.SelectedExportFolder;
        window.EnableSubfolder = settings.EnableSubfolder;
        if (window.SubfolderNameTextBox is not null)
            window.SubfolderNameTextBox.Text = settings.SubfolderName;
        window.FixedFolderPath = settings.FixedFolderPath;
        
        window.SelectedProcessingMethod = settings.SelectedProcessingMethod;
        window.OptimizeDxf = settings.OptimizeDxf;
        
        window.EnableFileNameConstructor = settings.EnableFileNameConstructor;
        window.TokenService.FileNameTemplate = settings.FileNameTemplate;
        
        if (window.SettingsExpander is not null)
            window.SettingsExpander.IsExpanded = settings.IsExpanded;
        
        window.TemplatePresets.Clear();
        foreach (var presetData in settings.TemplatePresets)
        {
            window.TemplatePresets.Add(new TemplatePreset
            {
                Name = presetData.Name,
                Template = presetData.Template
            });
        }
        
        if (!string.IsNullOrEmpty(settings.SelectedTemplatePresetName))
        {
            window.SelectedTemplatePreset = window.TemplatePresets
                .FirstOrDefault(p => p.Name == settings.SelectedTemplatePresetName);
        }

        window.CustomPropertiesList.Clear();
        settings.CustomProperties.ForEach(window.CustomPropertiesList.Add);

        foreach (var columnName in settings.ColumnOrder)
        {
            var presetProperty = window.PresetIProperties.FirstOrDefault(p => p.ColumnHeader == columnName);
            if (presetProperty is not null)
            {
                window.AddIPropertyColumn(presetProperty);
            }
            else
            {
                window.AddCustomIPropertyColumn(columnName);
            }
        }

        foreach (var settingData in settings.LayerSettings)
        {
            var layerSetting = window.LayerSettings.FirstOrDefault(ls => ls.DisplayName == settingData.DisplayName);
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