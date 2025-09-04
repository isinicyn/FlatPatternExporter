using System.Collections.ObjectModel;
using System.Collections.Specialized;
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
                    // Добавляем колонку в DataGrid
                    _mainWindow.AddIPropertyColumn(selectedProperty);
                    
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

                // Проверяем, нет ли уже столбца с таким заголовком
                if (_mainWindow.PartsDataGrid.Columns.Any(c => c.Header as string == columnHeader))
                {
                    MessageBox.Show($"Столбец с именем '{columnHeader}' уже существует.", "Предупреждение",
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
                        IsAdded = true // Сразу помечаем как добавленное
                    };
                    
                    PresetIProperties.Add(newUserProperty);

                    // Отключаем интерфейс Inventor на время операции
                    _mainWindow.SetInventorUserInterfaceState(true);

                    try
                    {
                        // Создаем колонку (данные заполнятся автоматически)
                        _mainWindow.AddUserDefinedIPropertyColumn(userDefinedPropertyName);
                    }
                    finally
                    {
                        _mainWindow.SetInventorUserInterfaceState(false);
                    }
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
                        Category = userProperty.Category
                    };

                    PresetIProperties.Add(presetProperty);
                }
            }
        }

        private void RemovePropertyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is PresetIProperty property)
            {
                // Находим колонку по заголовку
                var columnToRemove = _mainWindow.PartsDataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == property.ColumnHeader);
                if (columnToRemove != null)
                {
                    // Удаляем колонку из DataGrid
                    _mainWindow.PartsDataGrid.Columns.Remove(columnToRemove);
                    
                    // Проверяем, является ли это пользовательским свойством
                    var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => 
                        p.ColumnHeader == property.ColumnHeader || p.InventorPropertyName == property.InventorPropertyName);
                    if (userProperty != null)
                    {
                        // Удаляем из реестра по InternalName
                        PropertyMetadataRegistry.RemoveUserDefinedProperty(userProperty.InternalName);
                        
                        // Удаляем из основной коллекции
                        PresetIProperties.Remove(property);
                    }
                    
                    // Обновляем состояние IsAdded
                    property.IsAdded = false;
                }
            }
        }
}