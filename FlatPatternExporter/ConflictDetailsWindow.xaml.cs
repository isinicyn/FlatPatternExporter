using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;



public class ConflictPartNumberGroup
{
    public string PartNumber { get; set; } = string.Empty;
    public List<ConflictFileInfo> Files { get; set; } = new();
}

public class ConflictFileInfo
{
    public string FileName { get; set; } = string.Empty;
    public string ModelState { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
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
            }).ToList()
        }).ToList();
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
        if (sender is MenuItem menuItem && menuItem.DataContext is ConflictFileInfo fileInfo)
        {
            ExecuteOpenFile(fileInfo);
        }
    }

    // Обработчик кнопки "ОК"
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}