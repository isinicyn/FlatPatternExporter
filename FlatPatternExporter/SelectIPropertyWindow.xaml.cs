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
    private readonly ObservableCollection<PresetIProperty> _presetIProperties;
    private bool _isUpdatingAvailableProperties; // Flag to prevent premature closure

        public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties, FlatPatternExporterMainWindow mainWindow)
        {
            InitializeComponent();

            _presetIProperties = presetIProperties;
            _mainWindow = mainWindow; // Сохраняем ссылку на MainWindow

        // Инициализируем AvailableProperties со всеми свойствами
        AvailableProperties = new ObservableCollection<PresetIProperty>(_presetIProperties);
        DataContext = this;
        SelectedProperties = [];

            // Подписка на событие изменения коллекции AvailableProperties
            AvailableProperties.CollectionChanged += AvailableProperties_CollectionChanged;

            // Подписка на событие изменения коллекции PresetIProperties
            _presetIProperties.CollectionChanged += PresetIProperties_CollectionChanged;
            
            // Инициализируем состояние IsAdded для уже добавленных колонок
            UpdateAvailableProperties();
        }

    public ObservableCollection<PresetIProperty> AvailableProperties { get; set; }
    public List<PresetIProperty> SelectedProperties { get; }

        // Обработчик события изменения коллекции AvailableProperties
        private void AvailableProperties_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (!_isUpdatingAvailableProperties && AvailableProperties.Count == 0)
            {
            // Закрываем окно без предупреждения
            Close();
            }
        }

        // Обработчик события изменения коллекции PresetIProperties
        private void PresetIProperties_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            // Обновляем список AvailableProperties при изменении коллекции PresetIProperties
            UpdateAvailableProperties();
        }

        // Обновление списка доступных свойств
        public void UpdateAvailableProperties()
        {
            try
            {
                _isUpdatingAvailableProperties = true;
                
                // Обновляем состояние IsAdded для всех свойств
                foreach (var property in AvailableProperties)
                {
                    property.IsAdded = _mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == property.ColumnHeader);
                }
            }
            finally
            {
                _isUpdatingAvailableProperties = false;
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
                // Проверяем, не добавлено ли уже это свойство
                if (_mainWindow._userDefinedPropertiesList.Contains(userDefinedPropertyName))
                {
                MessageBox.Show($"Свойство '{userDefinedPropertyName}' уже добавлено.", "Предупреждение",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

            if (_mainWindow.PartsDataGrid.Columns.Any(c => c.Header as string == userDefinedPropertyName))
                {
                MessageBox.Show($"Столбец с именем '{userDefinedPropertyName}' уже существует.", "Предупреждение",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Добавляем в список user defined properties
                _mainWindow._userDefinedPropertiesList.Add(userDefinedPropertyName);

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

            // Очищаем текстовое поле
            UserDefinedPropertyTextBox.Text = "";
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
                    
                    // Обновляем состояние IsAdded
                    property.IsAdded = false;
                }
            }
        }
}