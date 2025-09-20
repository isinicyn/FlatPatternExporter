using System.Collections.ObjectModel;
using System.ComponentModel;
using FlatPatternExporter.UI.Windows;

namespace FlatPatternExporter.Core;

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
            if (!UserDefinedProperties.Any(p => p.ColumnHeader == userProperty.ColumnHeader))
            {
                var presetProperty = new PresetIProperty
                {
                    ColumnHeader = userProperty.ColumnHeader,
                    ListDisplayName = userProperty.DisplayName,
                    InventorPropertyName = userProperty.InternalName,
                    Category = userProperty.Category,
                    IsUserDefined = true
                };

                UserDefinedProperties.Add(presetProperty);
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
                ColumnHeader = userProperty.ColumnHeader,
                ListDisplayName = userProperty.DisplayName,
                InventorPropertyName = userProperty.InventorPropertyName ?? userDefinedPropertyName,
                Category = userProperty.Category,
                IsAdded = false,
                IsUserDefined = true
            };

            UserDefinedProperties.Add(newUserProperty);
            return newUserProperty;
        }

        return null;
    }

    public bool RemoveUserDefinedProperty(PresetIProperty property)
    {
        return UserDefinedProperties.Remove(property);
    }
}
