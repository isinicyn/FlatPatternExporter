using System.ComponentModel;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Models;

public class PresetIProperty : INotifyPropertyChanged
{
    private bool _isAdded;

    public string InventorPropertyName { get; set; } = string.Empty;
    public bool IsUserDefined { get; set; } = false;

    public string ColumnHeader
    {
        get
        {
            var prop = PropertyMetadataRegistry.GetPropertyByInternalName(InventorPropertyName);
            return prop?.ColumnHeader ?? string.Empty;
        }
    }

    public string ListDisplayName
    {
        get
        {
            var prop = PropertyMetadataRegistry.GetPropertyByInternalName(InventorPropertyName);
            return prop?.DisplayName ?? string.Empty;
        }
    }

    public string Category
    {
        get
        {
            var prop = PropertyMetadataRegistry.GetPropertyByInternalName(InventorPropertyName);
            return prop?.Category ?? string.Empty;
        }
    }

    public bool IsAdded
    {
        get => _isAdded;
        set
        {
            if (_isAdded != value)
            {
                _isAdded = value;
                OnPropertyChanged(nameof(IsAdded));
            }
        }
    }

    public PresetIProperty()
    {
        LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
    }

    private void OnLanguageChanged(object? sender, EventArgs e)
    {
        OnPropertyChanged(nameof(ColumnHeader));
        OnPropertyChanged(nameof(ListDisplayName));
        OnPropertyChanged(nameof(Category));
    }

    protected void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}