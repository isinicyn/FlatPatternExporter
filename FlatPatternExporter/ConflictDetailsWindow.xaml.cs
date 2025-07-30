using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;

public partial class ConflictDetailsWindow
{
    private string? _selectedFilePath;

    // Конструктор принимает данные о конфликтах и заполняет TreeView
    public ConflictDetailsWindow(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        InitializeComponent();
        PopulateTreeView(conflictFileDetails);
    }

    // Заполняем TreeView данными о конфликтах
    private void PopulateTreeView(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        foreach (var entry in conflictFileDetails)
        {
            // Узел для обозначения детали
            var partNumberItem = new TreeViewItem
            {
                Header = $"Обозначение: {entry.Key}",
                IsExpanded = true,
                Tag = null // На уровне обозначений не указываем файл
            };

            // Файлы и состояния модели как дочерние элементы
            foreach (var conflictInfo in entry.Value)
            {
                var fileItem = new TreeViewItem
                {
                    Header = $"Файл: {conflictInfo.FileName} | Состояние: {conflictInfo.ModelState}",
                    Tag = conflictInfo.FileName // Сохраняем путь к файлу в Tag для дальнейшего использования
                };
                partNumberItem.Items.Add(fileItem);
            }

            ConflictTreeView.Items.Add(partNumberItem);
        }
    }

    // Обработчик для правого клика мыши
    private void ConflictTreeView_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var treeViewItem = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource);

        if (treeViewItem != null && treeViewItem.Tag != null) // Только для файлов
        {
            treeViewItem.Focus(); // Фокусируемся на нужном элементе
            _selectedFilePath = treeViewItem.Tag?.ToString();

            var contextMenu = new ContextMenu();
            var openFileMenuItem = new MenuItem { Header = "Открыть файл" };
            openFileMenuItem.Click += OpenFileMenuItem_Click;
            contextMenu.Items.Add(openFileMenuItem);
            contextMenu.IsOpen = true;
        }

        e.Handled = true; // Прерываем дальнейшую обработку события
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