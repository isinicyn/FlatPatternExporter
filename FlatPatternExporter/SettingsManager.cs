using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;

namespace FlatPatternExporter;

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
    public string FixedFolderPath { get; set; } = string.Empty;
    
    public ProcessingMethod SelectedProcessingMethod { get; set; } = ProcessingMethod.BOM;
    
    public bool IsPrimaryModelState { get; set; } = true;
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
            SelectedProcessingMethod = window.SelectedProcessingMethod,
            IsPrimaryModelState = window.IsPrimaryModelState,
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
        window.SelectedProcessingMethod = settings.SelectedProcessingMethod;
        window.IsPrimaryModelState = settings.IsPrimaryModelState;
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

        // Обновляем отображение фиксированной папки
        if (!string.IsNullOrEmpty(settings.FixedFolderPath))
        {
            window.UpdateFixedFolderPathDisplay(settings.FixedFolderPath);
        }
    }

    private static bool IsDefaultColumn(string columnName)
    {
        var defaultColumns = new[]
        {
            "ID", "Обозначение", "Наименование", "Состояние модели", 
            "Материал", "Толщина", "Кол.", "Изобр. детали", 
            "Изобр. развертки", "Обр."
        };
        return defaultColumns.Contains(columnName);
    }
}