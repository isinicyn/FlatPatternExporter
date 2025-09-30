using System.ComponentModel;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Models;

public class PresetIProperty : INotifyPropertyChanged
{
    private bool _isAdded;
    private string _substitutionValue = string.Empty;

    public string InventorPropertyName { get; set; } = string.Empty;
    public bool IsUserDefined { get; set; } = false;

    public string SubstitutionValue
    {
        get => _substitutionValue;
        set
        {
            if (_substitutionValue != value)
            {
                _substitutionValue = value;
                OnPropertyChanged(nameof(SubstitutionValue));

                if (string.IsNullOrEmpty(value))
                {
                    PropertyMetadataRegistry.PropertySubstitutions.Remove(InventorPropertyName);
                }
                else
                {
                    PropertyMetadataRegistry.PropertySubstitutions[InventorPropertyName] = value;
                }
            }
        }
    }

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