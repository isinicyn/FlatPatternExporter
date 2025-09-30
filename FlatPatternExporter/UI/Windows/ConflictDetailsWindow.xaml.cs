using System.Windows;
using System.Windows.Controls;
using FlatPatternExporter.Core;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.UI.Windows;


public partial class ConflictDetailsWindow
{
    private readonly Action<string, string>? _openDocumentAction;
    public List<ConflictPartNumberGroup> ConflictGroups { get; set; } = [];

    public ConflictDetailsWindow(Dictionary<string, List<PartConflictInfo>> conflictFileDetails,
                               Action<string, string>? openDocumentAction = null)
    {
        InitializeComponent();
        _openDocumentAction = openDocumentAction;
        ConflictGroups = ConflictDataProcessor.PrepareConflictData(conflictFileDetails);
        DataContext = this;
    }

    private void ExecuteOpenFile(ConflictFileInfo? fileInfo)
    {
        if (fileInfo?.FilePath == null)
            return;

        _openDocumentAction?.Invoke(fileInfo.FilePath, fileInfo.ModelState);
    }

    private void OpenFileMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem)
        {
            // Find the DataGrid that contains this menu
            var dataGrid = menuItem.GetValue(FrameworkElement.DataContextProperty) as DataGrid;
            if (dataGrid?.SelectedItem is ConflictFileInfo fileInfo)
            {
                ExecuteOpenFile(fileInfo);
            }
            // Try to find via PlacementTarget
            else if (menuItem.Parent is ContextMenu contextMenu && 
                     contextMenu.PlacementTarget is DataGrid targetGrid &&
                     targetGrid.SelectedItem is ConflictFileInfo selectedFile)
            {
                ExecuteOpenFile(selectedFile);
            }
        }
    }

    private void DataGrid_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is DataGrid dataGrid)
        {
            dataGrid.UnselectAll();
        }
    }

    // "OK" button handler
    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}