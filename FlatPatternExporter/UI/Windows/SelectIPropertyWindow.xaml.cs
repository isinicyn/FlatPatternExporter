using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Data;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter.UI.Windows;
public partial class SelectIPropertyWindow : CustomChromeWindow
{
    private readonly PropertyListManager _propertyListManager;
    private readonly LocalizationManager _localizationManager = LocalizationManager.Instance;
    private readonly Func<string, bool> _isPropertyAdded;
    private readonly Action<PresetIProperty> _addIPropertyColumn;
    private readonly Action<string> _addUserDefinedColumn;
    private readonly Action<string, bool> _removeProperty;
    private bool _isInitialized = false;

    public SelectIPropertyWindow(
        ObservableCollection<PresetIProperty> presetIProperties,
        Func<string, bool> isPropertyAdded,
        Action<PresetIProperty> addIPropertyColumn,
        Action<string> addUserDefinedColumn,
        Action<string, bool> removeProperty)
    {
        InitializeComponent();

        _propertyListManager = new PropertyListManager(presetIProperties);
        _isPropertyAdded = isPropertyAdded;
        _addIPropertyColumn = addIPropertyColumn;
        _addUserDefinedColumn = addUserDefinedColumn;
        _removeProperty = removeProperty;

        DataContext = this;

        _propertyListManager.InitializeProperties();
        InitializeFilters();
        UpdatePropertyStates();

        _isInitialized = true;
    }

    public ObservableCollection<PresetIProperty> StandardProperties => _propertyListManager.StandardProperties;
    public ObservableCollection<PresetIProperty> UserDefinedProperties => _propertyListManager.UserDefinedProperties;

    private void InitializeFilters()
    {
        var categories = _propertyListManager.GetCategories();
        foreach (var category in categories)
        {
            CategoryFilter.Items.Add(new ComboBoxItem { Content = category });
        }
    }

    public void UpdatePropertyStates()
    {
        foreach (var property in StandardProperties.Concat(UserDefinedProperties))
        {
            property.IsAdded = _isPropertyAdded(property.InventorPropertyName);
        }
    }


    private void StandardSearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        string searchFilter = StandardSearchBox.Text;
        ClearStandardSearchButton.IsEnabled = !string.IsNullOrEmpty(searchFilter);
        _propertyListManager.SetSearchFilter(searchFilter);
        ApplyFilters();
    }

    private void ClearStandardSearchButton_Click(object sender, RoutedEventArgs e)
    {
        StandardSearchBox.Text = string.Empty;
        ClearStandardSearchButton.IsEnabled = false;
        StandardSearchBox.Focus();
    }

    private void CategoryFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_isInitialized) return;

        string categoryFilter = "";
        if (CategoryFilter.SelectedIndex != 0 && CategoryFilter.SelectedItem is ComboBoxItem item)
        {
            categoryFilter = item.Content?.ToString() ?? "";
        }
        _propertyListManager.SetCategoryFilter(categoryFilter);
        ApplyFilters();
    }

    private void ApplyFilters()
    {
        var view = CollectionViewSource.GetDefaultView(StandardPropertiesListBox.ItemsSource);
        if (view != null)
        {
            _propertyListManager.ApplyFilters(view);
        }
    }

    private void StandardPropertiesListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (StandardPropertiesListBox.SelectedItem is PresetIProperty selectedProperty && !selectedProperty.IsAdded)
        {
            AddPropertyToGrid(selectedProperty);
        }
    }

    private void UserPropertiesListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (UserPropertiesListBox.SelectedItem is PresetIProperty selectedProperty && !selectedProperty.IsAdded)
        {
            AddUserPropertyToGrid(selectedProperty);
        }
    }

    private void AddPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is PresetIProperty property)
        {
            AddPropertyToGrid(property);
        }
    }

    private void RemovePropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is PresetIProperty property)
        {
            RemovePropertyFromGrid(property);
        }
    }

    private void AddUserPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is PresetIProperty property)
        {
            AddUserPropertyToGrid(property);
        }
    }

    private void RemoveUserPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is PresetIProperty property)
        {
            RemovePropertyFromGrid(property);
        }
    }

    private void DeleteUserPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is PresetIProperty property)
        {
            var result = MessageBox.Show(
                _localizationManager.GetString("Confirm_DeleteProperty", property.ListDisplayName),
                _localizationManager.GetString("Confirm_DeletePropertyTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
                
            if (result == MessageBoxResult.Yes)
            {
                _removeProperty(property.InventorPropertyName, false);

                if (!string.IsNullOrEmpty(property.InventorPropertyName))
                {
                    PropertyMetadataRegistry.RemoveUserDefinedProperty(property.InventorPropertyName);
                }

                _propertyListManager.RemoveUserDefinedProperty(property);
            }
        }
    }

    private void AddPropertyToGrid(PresetIProperty property)
    {
        if (!property.IsAdded)
        {
            _addIPropertyColumn(property);
            property.IsAdded = true;
        }
    }

    private void AddUserPropertyToGrid(PresetIProperty property)
    {
        if (!property.IsAdded)
        {
            var internalName = property.InventorPropertyName;
            if (!string.IsNullOrEmpty(internalName) && PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
            {
                var propertyName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);
                if (!string.IsNullOrEmpty(propertyName))
                {
                    _addUserDefinedColumn(propertyName);
                }
            }
            property.IsAdded = true;
        }
    }

    private void RemovePropertyFromGrid(PresetIProperty property)
    {
        _removeProperty(property.InventorPropertyName, true);
        property.IsAdded = false;
    }

    private void UserDefinedPropertyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        AddUserDefinedPropertyButton.IsEnabled = !string.IsNullOrWhiteSpace(UserDefinedPropertyTextBox.Text);
    }

    private void UserDefinedPropertyTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Enter && AddUserDefinedPropertyButton.IsEnabled)
        {
            AddUserDefinedPropertyButton_Click(AddUserDefinedPropertyButton, new RoutedEventArgs());
        }
    }

    private void AddUserDefinedPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        string userDefinedPropertyName = UserDefinedPropertyTextBox.Text.Trim();

        var newProperty = _propertyListManager.CreateUserDefinedProperty(userDefinedPropertyName);
        if (newProperty == null)
        {
            MessageBox.Show(_localizationManager.GetString("Warning_PropertyAlreadyAdded", userDefinedPropertyName), _localizationManager.GetString("MessageBox_Warning"),
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        UserDefinedPropertyTextBox.Text = "";
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}