using System.ComponentModel;

namespace FlatPatternExporter.Models;

public class PresetIProperty : INotifyPropertyChanged
{
    private bool _isAdded;

    public string ColumnHeader { get; set; } = string.Empty;
    public string ListDisplayName { get; set; } = string.Empty;
    public string InventorPropertyName { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public bool IsUserDefined { get; set; } = false;

    public bool IsAdded
    {
        get => _isAdded;
        set
        {
            if (_isAdded != value)
            {
                _isAdded = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsAdded)));
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}