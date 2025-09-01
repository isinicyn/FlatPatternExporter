using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;



public class ConflictPartNumberGroup
{
    public string PartNumber { get; set; } = "";
    public List<ConflictFileInfo> Files { get; set; } = new();
}

public class ConflictFileInfo
{
    public string FileName { get; set; } = "";
    public string ModelState { get; set; } = "";
    public string FilePath { get; set; } = "";
    public string FileNameOnly => Path.GetFileName(FileName);
}

public partial class ConflictDetailsWindow
{
    public List<ConflictPartNumberGroup> ConflictGroups { get; set; } = new();

    public ConflictDetailsWindow(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        InitializeComponent();
        PrepareConflictData(conflictFileDetails);
        DataContext = this;
    }

    private void PrepareConflictData(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        ConflictGroups = conflictFileDetails.Select(entry => new ConflictPartNumberGroup
        {
            PartNumber = entry.Key,
            Files = entry.Value.Select(conflictInfo => new ConflictFileInfo
            {
                FileName = conflictInfo.FileName,
                ModelState = conflictInfo.ModelState,
                FilePath = conflictInfo.FileName
            }).OrderBy(f => f.FileName).ToList()
        }).OrderBy(g => g.PartNumber).ToList();
    }

    private void ExecuteOpenFile(ConflictFileInfo? fileInfo)
    {
        if (fileInfo?.FilePath == null)
            return;

        // Получаем ссылку на главное окно для использования централизованного метода
        var mainWindow = Owner as FlatPatternExporterMainWindow;
        if (mainWindow != null)
        {
            // Используем централизованный метод открытия с состоянием модели
            // Обработка ошибок уже есть внутри OpenInventorDocument()
            mainWindow.OpenInventorDocument(fileInfo.FilePath, fileInfo.ModelState);
        }
    }

    private void OpenFileMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem)
        {
            // Находим DataGrid, который содержит данное меню
            var dataGrid = menuItem.GetValue(FrameworkElement.DataContextProperty) as DataGrid;
            if (dataGrid?.SelectedItem is ConflictFileInfo fileInfo)
            {
                ExecuteOpenFile(fileInfo);
            }
            // Попробуем найти через PlacementTarget
            else if (menuItem.Parent is ContextMenu contextMenu && 
                     contextMenu.PlacementTarget is DataGrid targetGrid &&
                     targetGrid.SelectedItem is ConflictFileInfo selectedFile)
            {
                ExecuteOpenFile(selectedFile);
            }
        }
    }

    // Обработчик кнопки "ОК"
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}