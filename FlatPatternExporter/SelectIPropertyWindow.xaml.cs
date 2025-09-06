using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;
public partial class SelectIPropertyWindow : Window
{
    private readonly FlatPatternExporterMainWindow _mainWindow; // Поле для хранения ссылки на MainWindow

        public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties, FlatPatternExporterMainWindow mainWindow)
        {
            InitializeComponent();

            PresetIProperties = presetIProperties;
            _mainWindow = mainWindow; // Сохраняем ссылку на MainWindow

        DataContext = this;
        SelectedProperties = [];

            // Инициализируем пользовательские свойства из PropertyMetadataRegistry
            InitializeUserDefinedProperties();

            // Инициализируем состояние IsAdded для уже добавленных колонок
            UpdatePropertyStates();
        }

    public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }
    public List<PresetIProperty> SelectedProperties { get; }

        // Обновление состояния IsAdded для всех свойств
        public void UpdatePropertyStates()
        {
            // Обновляем состояние IsAdded для всех свойств
            foreach (var property in PresetIProperties)
            {
                property.IsAdded = _mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == property.ColumnHeader);
            }
        }



        private void iPropertyListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (IPropertyListBox.SelectedItem is PresetIProperty selectedProperty)
            {
                // Проверяем, есть ли уже такая колонка
                if (!_mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == selectedProperty.ColumnHeader))
                {
                    // Проверяем, является ли это UDP свойством
                    var internalName = PropertyMetadataRegistry.GetInternalNameByColumnHeader(selectedProperty.ColumnHeader);
                    if (!string.IsNullOrEmpty(internalName) && PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
                    {
                        // Извлекаем имя свойства из внутреннего имени
                        var propertyName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);
                        if (!string.IsNullOrEmpty(propertyName))
                        {
                            _mainWindow.AddUserDefinedIPropertyColumn(propertyName);
                        }
                    }
                    else
                    {
                        // Обычное свойство
                        _mainWindow.AddIPropertyColumn(selectedProperty);
                    }
                    
                    // Обновляем состояние IsAdded
                    selectedProperty.IsAdded = true;
                }
            }
        }



        private void UserDefinedPropertyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Активируем кнопку '+' только если текстовое поле не пустое
            AddUserDefinedPropertyButton.IsEnabled = !string.IsNullOrWhiteSpace(UserDefinedPropertyTextBox.Text);
        }

        private void AddUserDefinedPropertyButton_Click(object sender, RoutedEventArgs e)
        {
            string userDefinedPropertyName = UserDefinedPropertyTextBox.Text;
            
            if (!string.IsNullOrWhiteSpace(userDefinedPropertyName))
            {
                // Генерируем internal name и column header с префиксами
                var internalName = $"UDP_{userDefinedPropertyName}";
                var columnHeader = $"(Пользов.) {userDefinedPropertyName}";

                // Проверяем, не добавлено ли уже это свойство по InternalName
                if (PropertyMetadataRegistry.UserDefinedProperties.Any(p => p.InternalName == internalName))
                {
                    MessageBox.Show($"Свойство '{userDefinedPropertyName}' уже добавлено.", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Добавляем в реестр пользовательских свойств
                PropertyMetadataRegistry.AddUserDefinedProperty(userDefinedPropertyName);

                // Получаем созданное определение свойства
                var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
                if (userProperty != null)
                {
                    // Создаем PresetIProperty и добавляем в основную коллекцию
                    var newUserProperty = new PresetIProperty
                    {
                        ColumnHeader = userProperty.ColumnHeader,
                        ListDisplayName = userProperty.DisplayName,
                        InventorPropertyName = userProperty.InventorPropertyName ?? userDefinedPropertyName,
                        Category = userProperty.Category,
                        IsAdded = false, // Не помечаем как добавленное, так как колонка не создается
                        IsUserDefined = true // Помечаем как пользовательское свойство
                    };
                    
                    PresetIProperties.Add(newUserProperty);
                }

                // Очищаем текстовое поле
                UserDefinedPropertyTextBox.Text = "";
            }
        }

        /// <summary>
        /// Инициализирует пользовательские свойства из PropertyMetadataRegistry
        /// </summary>
        private void InitializeUserDefinedProperties()
        {
            foreach (var userProperty in PropertyMetadataRegistry.UserDefinedProperties)
            {
                // Проверяем, нет ли уже этого свойства в основной коллекции по ColumnHeader
                if (!PresetIProperties.Any(p => p.ColumnHeader == userProperty.ColumnHeader))
                {
                    var presetProperty = new PresetIProperty
                    {
                        ColumnHeader = userProperty.ColumnHeader,
                        ListDisplayName = userProperty.DisplayName,
                        InventorPropertyName = userProperty.InternalName,
                        Category = userProperty.Category,
                        IsUserDefined = true
                    };

                    PresetIProperties.Add(presetProperty);
                }
            }
        }

        private void RemovePropertyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is PresetIProperty property)
            {
                if (property.IsUserDefined)
                {
                    // Для пользовательских свойств показываем контекстное меню
                    button.ContextMenu.PlacementTarget = button;
                    button.ContextMenu.IsOpen = true;
                }
                else
                {
                    // Для обычных свойств просто удаляем колонку
                    RemoveColumn(property);
                }
            }
        }
        
        private void RemoveColumnOnlyMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is PresetIProperty property)
            {
                // Удаляем только колонку, данные UDP остаются
                _mainWindow.RemoveDataGridColumn(property.ColumnHeader, removeDataOnly: true);
                property.IsAdded = false;
            }
        }
        
        private void RemoveColumnAndPropertyMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is PresetIProperty property)
            {
                // Полное удаление колонки и UDP свойства
                _mainWindow.RemoveDataGridColumn(property.ColumnHeader, removeDataOnly: false);
                
                // Удаляем из основной коллекции
                PresetIProperties.Remove(property);
            }
        }
        
        private void RemoveColumn(PresetIProperty property)
        {
            // Удаляем только колонку для обычных свойств
            _mainWindow.RemoveDataGridColumn(property.ColumnHeader, removeDataOnly: true);
            property.IsAdded = false;
        }
}