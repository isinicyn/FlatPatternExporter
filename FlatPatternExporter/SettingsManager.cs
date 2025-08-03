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
    public bool IsVisible { get; set; } = true;
    public string CustomName { get; set; } = string.Empty;
    public string SelectedColor { get; set; } = "White";
    public string SelectedLineType { get; set; } = "Default";
    public string OriginalName { get; set; } = string.Empty;
}

[Serializable]
public class ApplicationSettings
{
    public ObservableCollection<string> ActiveColumns { get; set; } = new();
    public ObservableCollection<string> CustomProperties { get; set; } = new();
    
    public bool ExcludeReferenceParts { get; set; } = true;
    public bool ExcludePurchasedParts { get; set; } = true;
    public bool IncludeLibraryComponents { get; set; } = false;
    
    public bool OrganizeByMaterial { get; set; } = false;
    public bool OrganizeByThickness { get; set; } = false;
    public bool IncludeQuantityInFileName { get; set; } = false;
    
    public bool EnableSplineReplacement { get; set; } = false;
    public int SelectedSplineReplacementIndex { get; set; } = 0;
    public int SelectedAcadVersionIndex { get; set; } = 5;
    public bool MergeProfilesIntoPolyline { get; set; } = true;
    public bool RebaseGeometry { get; set; } = true;
    public bool TrimCenterlines { get; set; } = false;
    
    public ExportFolderType SelectedExportFolder { get; set; } = ExportFolderType.ChooseFolder;
    public bool EnableSubfolder { get; set; } = false;
    public string SubfolderName { get; set; } = string.Empty;
    public string FixedFolderPath { get; set; } = string.Empty;
    
    public ProcessingMethod SelectedProcessingMethod { get; set; } = ProcessingMethod.BOM;
    
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
            IncludeLibraryComponents = window.IncludeLibraryComponents,
            OrganizeByMaterial = window.OrganizeByMaterial,
            OrganizeByThickness = window.OrganizeByThickness,
            IncludeQuantityInFileName = window.IncludeQuantityInFileName,
            EnableSplineReplacement = window.EnableSplineReplacement,
            SelectedSplineReplacementIndex = window.SelectedSplineReplacementIndex,
            SelectedAcadVersionIndex = window.SelectedAcadVersionIndex,
            MergeProfilesIntoPolyline = window.MergeProfilesIntoPolyline,
            RebaseGeometry = window.RebaseGeometry,
            TrimCenterlines = window.TrimCenterlines,
            SelectedExportFolder = window.SelectedExportFolder,
            EnableSubfolder = window.EnableSubfolder,
            SubfolderName = window.SubfolderNameTextBox?.Text ?? string.Empty,
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            FixedFolderPath = window.FixedFolderPath
        };

        foreach (var column in window.PartsDataGrid.Columns)
        {
            if (column.Header is string headerName && !IsDefaultColumn(headerName))
            {
                settings.ActiveColumns.Add(headerName);
            }
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
                IsVisible = layerSetting.IsVisible,
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
        window.IncludeLibraryComponents = settings.IncludeLibraryComponents;
        window.OrganizeByMaterial = settings.OrganizeByMaterial;
        window.OrganizeByThickness = settings.OrganizeByThickness;
        window.IncludeQuantityInFileName = settings.IncludeQuantityInFileName;
        window.EnableSplineReplacement = settings.EnableSplineReplacement;
        window.SelectedSplineReplacementIndex = settings.SelectedSplineReplacementIndex;
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

        window.CustomPropertiesList.Clear();
        foreach (var customProperty in settings.CustomProperties)
        {
            window.CustomPropertiesList.Add(customProperty);
        }

        foreach (var columnName in settings.ActiveColumns)
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
                layerSetting.IsVisible = settingData.IsVisible;
                layerSetting.CustomName = settingData.CustomName;
                layerSetting.SelectedColor = settingData.SelectedColor;
                layerSetting.SelectedLineType = settingData.SelectedLineType;
                layerSetting.OriginalName = settingData.OriginalName;
            }
        }
    }

    private static bool IsDefaultColumn(string columnName)
    {
        var defaultColumns = new[] { "Обр." };
        return defaultColumns.Contains(columnName);
    }
}