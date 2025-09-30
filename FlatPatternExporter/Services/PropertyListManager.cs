using System.Collections.ObjectModel;
using System.ComponentModel;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Services;

public class PropertyListManager(ObservableCollection<PresetIProperty> allProperties)
{
    private readonly ObservableCollection<PresetIProperty> _allProperties = allProperties;
    private string _searchFilter = "";
    private string _categoryFilter = "";

    public ObservableCollection<PresetIProperty> StandardProperties { get; } = [];
    public ObservableCollection<PresetIProperty> UserDefinedProperties { get; } = [];

    public void InitializeProperties()
    {
        foreach (var property in _allProperties)
        {
            if (property.IsUserDefined)
            {
                UserDefinedProperties.Add(property);
            }
            else
            {
                StandardProperties.Add(property);
            }
        }

        InitializeUserDefinedPropertiesFromRegistry();
    }

    public void InitializeUserDefinedPropertiesFromRegistry()
    {
        foreach (var userProperty in PropertyMetadataRegistry.UserDefinedProperties)
        {
            if (!UserDefinedProperties.Any(p => p.InventorPropertyName == userProperty.InternalName))
            {
                var existingProperty = _allProperties.FirstOrDefault(p => p.InventorPropertyName == userProperty.InternalName);

                if (existingProperty != null)
                {
                    UserDefinedProperties.Add(existingProperty);
                }
                else
                {
                    var presetProperty = new PresetIProperty
                    {
                        InventorPropertyName = userProperty.InternalName,
                        IsUserDefined = true
                    };

                    // Restore default value if exists
                    if (PropertyMetadataRegistry.PropertyDefaultValues.TryGetValue(userProperty.InternalName, out var defaultValue))
                    {
                        presetProperty.DefaultValue = defaultValue;
                    }

                    _allProperties.Add(presetProperty);
                    UserDefinedProperties.Add(presetProperty);
                }
            }
        }
    }

    public IEnumerable<string> GetCategories()
    {
        return StandardProperties.Select(p => p.Category).Distinct().OrderBy(c => c);
    }

    public void SetSearchFilter(string searchFilter)
    {
        _searchFilter = searchFilter;
    }

    public void SetCategoryFilter(string categoryFilter)
    {
        _categoryFilter = categoryFilter;
    }

    public void ApplyFilters(ICollectionView view)
    {
        view.Filter = item =>
        {
            if (item is PresetIProperty property)
            {
                bool matchesSearch = string.IsNullOrEmpty(_searchFilter) ||
                                    property.ListDisplayName.Contains(_searchFilter, StringComparison.OrdinalIgnoreCase);
                bool matchesCategory = string.IsNullOrEmpty(_categoryFilter) ||
                                      property.Category == _categoryFilter;
                return matchesSearch && matchesCategory;
            }
            return true;
        };
    }

    public PresetIProperty? CreateUserDefinedProperty(string userDefinedPropertyName)
    {
        if (string.IsNullOrWhiteSpace(userDefinedPropertyName))
            return null;

        var internalName = $"UDP_{userDefinedPropertyName}";

        if (PropertyMetadataRegistry.UserDefinedProperties.Any(p => p.InternalName == internalName))
            return null;

        PropertyMetadataRegistry.AddUserDefinedProperty(userDefinedPropertyName);

        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
        if (userProperty != null)
        {
            var newUserProperty = new PresetIProperty
            {
                InventorPropertyName = userProperty.InternalName,
                IsAdded = false,
                IsUserDefined = true
            };

            _allProperties.Add(newUserProperty);
            UserDefinedProperties.Add(newUserProperty);
            return newUserProperty;
        }

        return null;
    }

    public bool RemoveUserDefinedProperty(PresetIProperty property)
    {
        var removed = UserDefinedProperties.Remove(property);
        if (removed)
        {
            _allProperties.Remove(property);
        }
        return removed;
    }
}
