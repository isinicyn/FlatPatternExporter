using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Data;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;
public partial class SelectIPropertyWindow : Window
{
    private readonly FlatPatternExporterMainWindow _mainWindow;
    private readonly ObservableCollection<PresetIProperty> _allProperties;
    private string _searchFilter = "";
    private string _categoryFilter = "";
    private bool _isInitialized = false;

    public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties, FlatPatternExporterMainWindow mainWindow)
    {
        InitializeComponent();

        _mainWindow = mainWindow;
        _allProperties = presetIProperties;
        
        StandardProperties = [];
        UserDefinedProperties = [];
        
        DataContext = this;
        
        InitializeProperties();
        InitializeFilters();
        UpdatePropertyStates();
        
        _isInitialized = true;
    }

    public ObservableCollection<PresetIProperty> StandardProperties { get; set; }
    public ObservableCollection<PresetIProperty> UserDefinedProperties { get; set; }

    private void InitializeProperties()
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

    private void InitializeUserDefinedPropertiesFromRegistry()
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

    private void InitializeFilters()
    {
        var categories = StandardProperties.Select(p => p.Category).Distinct().OrderBy(c => c);
        foreach (var category in categories)
        {
            CategoryFilter.Items.Add(new ComboBoxItem { Content = category });
        }
    }

    public void UpdatePropertyStates()
    {
        foreach (var property in StandardProperties.Concat(UserDefinedProperties))
        {
            property.IsAdded = _mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == property.ColumnHeader);
        }
    }


    private void StandardSearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _searchFilter = StandardSearchBox.Text;
        ClearStandardSearchButton.IsEnabled = !string.IsNullOrEmpty(_searchFilter);
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
        
        if (CategoryFilter.SelectedIndex == 0)
        {
            _categoryFilter = "";
        }
        else if (CategoryFilter.SelectedItem is ComboBoxItem item)
        {
            _categoryFilter = item.Content?.ToString() ?? "";
        }
        ApplyFilters();
    }

    private void ApplyFilters()
    {
        var view = CollectionViewSource.GetDefaultView(StandardPropertiesListBox.ItemsSource);
        if (view != null)
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
                $"Вы уверены, что хотите удалить свойство '{property.ListDisplayName}' из реестра? Это действие нельзя отменить.",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
                
            if (result == MessageBoxResult.Yes)
            {
                _mainWindow.RemoveDataGridColumn(property.ColumnHeader, removeDataOnly: false);
                
                var internalName = PropertyMetadataRegistry.GetInternalNameByColumnHeader(property.ColumnHeader);
                if (!string.IsNullOrEmpty(internalName))
                {
                    PropertyMetadataRegistry.RemoveUserDefinedProperty(internalName);
                }
                
                UserDefinedProperties.Remove(property);
            }
        }
    }

    private void AddPropertyToGrid(PresetIProperty property)
    {
        if (!property.IsAdded)
        {
            _mainWindow.AddIPropertyColumn(property);
            property.IsAdded = true;
        }
    }

    private void AddUserPropertyToGrid(PresetIProperty property)
    {
        if (!property.IsAdded)
        {
            var internalName = PropertyMetadataRegistry.GetInternalNameByColumnHeader(property.ColumnHeader);
            if (!string.IsNullOrEmpty(internalName) && PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
            {
                var propertyName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);
                if (!string.IsNullOrEmpty(propertyName))
                {
                    _mainWindow.AddUserDefinedIPropertyColumn(propertyName);
                }
            }
            property.IsAdded = true;
        }
    }

    private void RemovePropertyFromGrid(PresetIProperty property)
    {
        _mainWindow.RemoveDataGridColumn(property.ColumnHeader, removeDataOnly: true);
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
        
        if (!string.IsNullOrWhiteSpace(userDefinedPropertyName))
        {
            var internalName = $"UDP_{userDefinedPropertyName}";
            var columnHeader = $"(Пользов.) {userDefinedPropertyName}";

            if (PropertyMetadataRegistry.UserDefinedProperties.Any(p => p.InternalName == internalName))
            {
                MessageBox.Show($"Свойство '{userDefinedPropertyName}' уже добавлено.", "Предупреждение",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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
            }

            UserDefinedPropertyTextBox.Text = "";
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}