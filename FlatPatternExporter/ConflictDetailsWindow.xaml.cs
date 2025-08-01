using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
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
    public string DisplayText => $"Файл: {FileName} | Состояние: {ModelState}";
    public string FilePath { get; set; } = string.Empty;
}

public partial class ConflictDetailsWindow
{
    private string? _selectedFilePath;
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
            PartNumber = $"Обозначение: {entry.Key}",
            Files = entry.Value.Select(conflictInfo => new ConflictFileInfo
            {
                FileName = conflictInfo.FileName,
                ModelState = conflictInfo.ModelState,
                FilePath = conflictInfo.FileName
            }).ToList()
        }).ToList();
    }

    private void ConflictTreeView_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var treeViewItem = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource);

        if (treeViewItem?.DataContext is ConflictFileInfo fileInfo)
        {
            treeViewItem.Focus();
            _selectedFilePath = fileInfo.FilePath;

            var contextMenu = new ContextMenu();
            var openFileMenuItem = new MenuItem { Header = "Открыть файл" };
            openFileMenuItem.Click += OpenFileMenuItem_Click;
            contextMenu.Items.Add(openFileMenuItem);
            contextMenu.IsOpen = true;
        }

        e.Handled = true;
    }

    // Поиск TreeViewItem для элемента
    private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
    {
        do
        {
            if (current is T) return (T)current;
            current = VisualTreeHelper.GetParent(current);
        } while (current != null);

        return null;
    }

    // Обработчик для пункта "Открыть файл"
    private void OpenFileMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_selectedFilePath) && File.Exists(_selectedFilePath))
            try
            {
                // Открываем файл
                Process.Start(new ProcessStartInfo
                {
                    FileName = _selectedFilePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть файл: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        else
            MessageBox.Show("Файл не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    // Обработчик кнопки "ОК"
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}