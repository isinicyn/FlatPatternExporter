using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;



public class ConflictItem
{
    public string PartNumber { get; set; } = "";
    public string FileName { get; set; } = "";
    public string ModelState { get; set; } = "";
    public string FilePath { get; set; } = "";
}

public partial class ConflictDetailsWindow
{
    public List<ConflictItem> ConflictItems { get; set; } = new();
    public ICollectionView ConflictItemsView { get; set; } = null!;

    public ConflictDetailsWindow(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        InitializeComponent();
        PrepareConflictData(conflictFileDetails);
        DataContext = this;
    }

    private void PrepareConflictData(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        ConflictItems = conflictFileDetails
            .SelectMany(entry => entry.Value.Select(conflictInfo => new ConflictItem
            {
                PartNumber = entry.Key,
                FileName = conflictInfo.FileName,
                ModelState = conflictInfo.ModelState,
                FilePath = conflictInfo.FileName
            }))
            .OrderBy(item => item.PartNumber)
            .ThenBy(item => item.FileName)
            .ToList();

        // Создаем CollectionView для группировки
        ConflictItemsView = CollectionViewSource.GetDefaultView(ConflictItems);
        ConflictItemsView.GroupDescriptions.Add(new PropertyGroupDescription("PartNumber"));
    }

    private void ExecuteOpenFile(ConflictItem? conflictItem)
    {
        if (conflictItem?.FilePath == null)
            return;

        // Получаем ссылку на главное окно для использования централизованного метода
        var mainWindow = Owner as FlatPatternExporterMainWindow;
        if (mainWindow != null)
        {
            // Используем централизованный метод открытия с состоянием модели
            // Обработка ошибок уже есть внутри OpenInventorDocument()
            mainWindow.OpenInventorDocument(conflictItem.FilePath, conflictItem.ModelState);
        }
    }

    private void OpenFileMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ConflictDataGrid.SelectedItem is ConflictItem conflictItem)
        {
            ExecuteOpenFile(conflictItem);
        }
    }

    // Обработчик кнопки "ОК"
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}