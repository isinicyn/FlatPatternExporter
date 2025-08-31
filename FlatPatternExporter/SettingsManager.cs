using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;

namespace FlatPatternExporter;

[Serializable]
public class LayerSettingData
{
    public string DisplayName { get; set; } = "";
    public bool IsChecked { get; set; }
    public string CustomName { get; set; } = "";
    public string SelectedColor { get; set; } = "White";
    public string SelectedLineType { get; set; } = "Default";
}

[Serializable]
public class TemplatePresetData
{
    public string Name { get; set; } = "";
    public string Template { get; set; } = "";
}

[Serializable]
public class ApplicationSettings
{
    public ObservableCollection<string> ColumnOrder { get; set; } = [];
    public ObservableCollection<string> CustomProperties { get; set; } = [];
    
    public bool ExcludeReferenceParts { get; set; } = true;
    public bool ExcludePurchasedParts { get; set; } = true;
    public bool ExcludePhantomParts { get; set; } = true;
    public bool IncludeLibraryComponents { get; set; }
    
    public bool OrganizeByMaterial { get; set; }
    public bool OrganizeByThickness { get; set; }
    
    public bool EnableSplineReplacement { get; set; }
    public int SelectedSplineReplacementIndex { get; set; } = 0;
    public string SplineTolerance { get; set; } = "0.01";
    public int SelectedAcadVersionIndex { get; set; } = 5;
    public bool MergeProfilesIntoPolyline { get; set; } = true;
    public bool RebaseGeometry { get; set; } = true;
    public bool TrimCenterlines { get; set; }
    
    public ExportFolderType SelectedExportFolder { get; set; } = ExportFolderType.ChooseFolder;
    public bool EnableSubfolder { get; set; }
    public string SubfolderName { get; set; } = "";
    public string FixedFolderPath { get; set; } = "";
    
    public ProcessingMethod SelectedProcessingMethod { get; set; } = ProcessingMethod.BOM;
    
    public bool OptimizeDxf { get; set; }
    
    public bool EnableFileNameConstructor { get; set; }
    public string FileNameTemplate { get; set; } = "{PartNumber}";
    
    public ObservableCollection<TemplatePresetData> TemplatePresets { get; set; } = [];
    public string SelectedTemplatePresetName { get; set; } = "";
    
    public ObservableCollection<LayerSettingData> LayerSettings { get; set; } = [];
    
    public bool IsExpanded { get; set; }
}

public static class SettingsManager
{
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "FlatPatternExporter",
        "settings.xml"
    );

    private static readonly XmlSerializer Serializer = new(typeof(ApplicationSettings));

    public static ApplicationSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return new ApplicationSettings();
            }

            using var fileStream = new FileStream(SettingsFilePath, FileMode.Open);
            return (ApplicationSettings?)Serializer.Deserialize(fileStream) ?? new ApplicationSettings();
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

            using var fileStream = new FileStream(SettingsFilePath, FileMode.Create);
            Serializer.Serialize(fileStream, settings);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error saving settings: {ex.Message}");
        }
    }

    public static ApplicationSettings CreateSettingsFromMainWindow(FlatPatternExporterMainWindow window)
    {
        var settings = new ApplicationSettings
        {
            ExcludeReferenceParts = window.ExcludeReferenceParts,
            ExcludePurchasedParts = window.ExcludePurchasedParts,
            ExcludePhantomParts = window.ExcludePhantomParts,
            IncludeLibraryComponents = window.IncludeLibraryComponents,
            OrganizeByMaterial = window.OrganizeByMaterial,
            OrganizeByThickness = window.OrganizeByThickness,
            EnableSplineReplacement = window.EnableSplineReplacement,
            SelectedSplineReplacementIndex = window.SelectedSplineReplacementIndex,
            SplineTolerance = window.SplineToleranceTextBox?.Text ?? "0.01",
            SelectedAcadVersionIndex = window.SelectedAcadVersionIndex,
            MergeProfilesIntoPolyline = window.MergeProfilesIntoPolyline,
            RebaseGeometry = window.RebaseGeometry,
            TrimCenterlines = window.TrimCenterlines,
            SelectedExportFolder = window.SelectedExportFolder,
            EnableSubfolder = window.EnableSubfolder,
            SubfolderName = window.SubfolderNameTextBox?.Text ?? "",
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            FixedFolderPath = window.FixedFolderPath,
            OptimizeDxf = window.OptimizeDxf,
            EnableFileNameConstructor = window.EnableFileNameConstructor,
            FileNameTemplate = window.TokenService.FileNameTemplate,
            IsExpanded = window.SettingsExpander?.IsExpanded ?? false
        };
        
        // Сохранение пресетов шаблонов
        settings.TemplatePresets.Clear();
        foreach (var preset in window.TemplatePresets)
        {
            settings.TemplatePresets.Add(new TemplatePresetData
            {
                Name = preset.Name,
                Template = preset.Template
            });
        }
        
        // Сохранение выбранного пресета
        settings.SelectedTemplatePresetName = window.SelectedTemplatePreset?.Name ?? "";

        var columnsInDisplayOrder = window.PartsDataGrid.Columns
            .Where(c => c.Header is string headerValue && !string.IsNullOrEmpty(headerValue))
            .OrderBy(c => c.DisplayIndex)
            .Select(c => (string)c.Header)
            .ToList();
            
        foreach (var columnName in columnsInDisplayOrder)
        {
            settings.ColumnOrder.Add(columnName);
        }

        foreach (var customProperty in window.CustomPropertiesList)
        {
            settings.CustomProperties.Add(customProperty);
        }

        var changedLayerSettings = window.LayerSettings.Where(ls => ls.HasChanges()).ToList();
        foreach (var layerSetting in changedLayerSettings)
        {
            settings.LayerSettings.Add(new LayerSettingData
            {
                DisplayName = layerSetting.DisplayName,
                IsChecked = layerSetting.IsChecked,
                CustomName = layerSetting.CustomName,
                SelectedColor = layerSetting.SelectedColor,
                SelectedLineType = layerSetting.SelectedLineType
            });
        }

        return settings;
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
        if (window.SplineToleranceTextBox != null)
            window.SplineToleranceTextBox.Text = settings.SplineTolerance;
        window.SelectedAcadVersionIndex = settings.SelectedAcadVersionIndex;
        window.MergeProfilesIntoPolyline = settings.MergeProfilesIntoPolyline;
        window.RebaseGeometry = settings.RebaseGeometry;
        window.TrimCenterlines = settings.TrimCenterlines;
        window.SelectedExportFolder = settings.SelectedExportFolder;
        window.EnableSubfolder = settings.EnableSubfolder;
        if (window.SubfolderNameTextBox != null)
            window.SubfolderNameTextBox.Text = settings.SubfolderName;
        window.SelectedProcessingMethod = settings.SelectedProcessingMethod;
        window.FixedFolderPath = settings.FixedFolderPath;
        window.OptimizeDxf = settings.OptimizeDxf;
        window.EnableFileNameConstructor = settings.EnableFileNameConstructor;
        window.TokenService.FileNameTemplate = settings.FileNameTemplate;
        
        if (window.SettingsExpander != null)
            window.SettingsExpander.IsExpanded = settings.IsExpanded;
        
        // Загрузка пресетов шаблонов
        window.TemplatePresets.Clear();
        foreach (var presetData in settings.TemplatePresets)
        {
            window.TemplatePresets.Add(new TemplatePreset
            {
                Name = presetData.Name,
                Template = presetData.Template
            });
        }
        
        // Восстановление выбранного пресета
        if (!string.IsNullOrEmpty(settings.SelectedTemplatePresetName))
        {
            window.SelectedTemplatePreset = window.TemplatePresets
                .FirstOrDefault(p => p.Name == settings.SelectedTemplatePresetName);
        }

        window.CustomPropertiesList.Clear();
        foreach (var customProperty in settings.CustomProperties)
        {
            window.CustomPropertiesList.Add(customProperty);
        }

        foreach (var columnName in settings.ColumnOrder)
        {
            var presetProperty = window.PresetIProperties.FirstOrDefault(p => p.ColumnHeader == columnName);
            if (presetProperty != null)
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
            if (layerSetting != null)
            {
                layerSetting.IsChecked = settingData.IsChecked;
                layerSetting.CustomName = settingData.CustomName;
                layerSetting.SelectedColor = settingData.SelectedColor;
                layerSetting.SelectedLineType = settingData.SelectedLineType;
            }
        }
    }

}