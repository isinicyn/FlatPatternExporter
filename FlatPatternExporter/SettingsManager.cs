using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;

namespace FlatPatternExporter;

[Serializable]
public class LayerSettingData
{
    public string DisplayName { get; set; } = string.Empty;
    public string LayerName { get; set; } = string.Empty;
    public bool HasVisibilityOption { get; set; } = true;
    public bool IsChecked { get; set; } = false;
    public string CustomName { get; set; } = string.Empty;
    public string SelectedColor { get; set; } = "White";
    public string SelectedLineType { get; set; } = "Default";
    public string OriginalName { get; set; } = string.Empty;
}

[Serializable]
public class TemplatePresetData
{
    public string Name { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
}

[Serializable]
public class ApplicationSettings
{
    public ObservableCollection<string> ColumnOrder { get; set; } = new();
    public ObservableCollection<string> CustomProperties { get; set; } = new();
    
    public bool ExcludeReferenceParts { get; set; } = true;
    public bool ExcludePurchasedParts { get; set; } = true;
    public bool ExcludePhantomParts { get; set; } = true;
    public bool IncludeLibraryComponents { get; set; } = false;
    
    public bool OrganizeByMaterial { get; set; } = false;
    public bool OrganizeByThickness { get; set; } = false;
    
    public bool EnableSplineReplacement { get; set; } = false;
    public int SelectedSplineReplacementIndex { get; set; } = 0;
    public string SplineTolerance { get; set; } = "0.01";
    public int SelectedAcadVersionIndex { get; set; } = 5;
    public bool MergeProfilesIntoPolyline { get; set; } = true;
    public bool RebaseGeometry { get; set; } = true;
    public bool TrimCenterlines { get; set; } = false;
    
    public ExportFolderType SelectedExportFolder { get; set; } = ExportFolderType.ChooseFolder;
    public bool EnableSubfolder { get; set; } = false;
    public string SubfolderName { get; set; } = string.Empty;
    public string FixedFolderPath { get; set; } = string.Empty;
    
    public ProcessingMethod SelectedProcessingMethod { get; set; } = ProcessingMethod.BOM;
    
    public bool EnableFileNameConstructor { get; set; } = false;
    public string FileNameTemplate { get; set; } = "{PartNumber}";
    
    public ObservableCollection<TemplatePresetData> TemplatePresets { get; set; } = new();
    public string SelectedTemplatePresetName { get; set; } = string.Empty;
    
    public ObservableCollection<LayerSettingData> LayerSettings { get; set; } = new();
}

public static class SettingsManager
{
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "FlatPatternExporter",
        "settings.xml"
    );

    public static ApplicationSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return new ApplicationSettings();
            }

            var serializer = new XmlSerializer(typeof(ApplicationSettings));
            using var fileStream = new FileStream(SettingsFilePath, FileMode.Open);
            return (ApplicationSettings?)serializer.Deserialize(fileStream) ?? new ApplicationSettings();
        }
        catch
        {
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

            var serializer = new XmlSerializer(typeof(ApplicationSettings));
            using var fileStream = new FileStream(SettingsFilePath, FileMode.Create);
            serializer.Serialize(fileStream, settings);
        }
        catch
        {
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
            SubfolderName = window.SubfolderNameTextBox?.Text ?? string.Empty,
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            FixedFolderPath = window.FixedFolderPath,
            EnableFileNameConstructor = window.EnableFileNameConstructor,
            FileNameTemplate = window.TokenService.FileNameTemplate
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
        settings.SelectedTemplatePresetName = window.SelectedTemplatePreset?.Name ?? string.Empty;

        var columnsInDisplayOrder = window.PartsDataGrid.Columns
            .Where(c => c.Header is string)
            .OrderBy(c => c.DisplayIndex)
            .Select(c => c.Header as string)
            .Where(name => !string.IsNullOrEmpty(name))
            .ToList();
            
        foreach (var columnName in columnsInDisplayOrder)
        {
            settings.ColumnOrder.Add(columnName!);
        }

        foreach (var customProperty in window.CustomPropertiesList)
        {
            settings.CustomProperties.Add(customProperty);
        }

        foreach (var layerSetting in window.LayerSettings)
        {
            settings.LayerSettings.Add(new LayerSettingData
            {
                DisplayName = layerSetting.DisplayName,
                LayerName = layerSetting.LayerName,
                HasVisibilityOption = layerSetting.HasVisibilityOption,
                IsChecked = layerSetting.IsChecked,
                CustomName = layerSetting.CustomName,
                SelectedColor = layerSetting.SelectedColor,
                SelectedLineType = layerSetting.SelectedLineType,
                OriginalName = layerSetting.OriginalName
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
        window.EnableFileNameConstructor = settings.EnableFileNameConstructor;
        window.TokenService.FileNameTemplate = settings.FileNameTemplate;
        
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
            var presetProperty = window.PresetIProperties.FirstOrDefault(p => p.InternalName == columnName);
            if (presetProperty != null)
            {
                window.AddIPropertyColumn(presetProperty);
            }
            else
            {
                window.AddCustomIPropertyColumn(columnName);
            }
        }

        if (settings.LayerSettings.Count > 0)
        {
            for (int i = 0; i < Math.Min(settings.LayerSettings.Count, window.LayerSettings.Count); i++)
            {
                var settingData = settings.LayerSettings[i];
                var layerSetting = window.LayerSettings[i];
                
                layerSetting.DisplayName = settingData.DisplayName;
                layerSetting.LayerName = settingData.LayerName;
                layerSetting.IsChecked = settingData.IsChecked;
                layerSetting.CustomName = settingData.CustomName;
                layerSetting.SelectedColor = settingData.SelectedColor;
                layerSetting.SelectedLineType = settingData.SelectedLineType;
                layerSetting.OriginalName = settingData.OriginalName;
            }
        }
    }

}