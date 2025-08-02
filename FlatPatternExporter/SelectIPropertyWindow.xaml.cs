using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Input;
using System.Windows.Controls;
using Application = System.Windows.Application;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using ListBox = System.Windows.Controls.ListBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;

namespace FlatPatternExporter
{
    public partial class SelectIPropertyWindow : Window
    {
        private bool _isMouseOverDataGrid; // Добавили переменную
        private readonly FlatPatternExporterMainWindow _mainWindow; // Поле для хранения ссылки на MainWindow
        private readonly ObservableCollection<PresetIProperty> _presetIProperties;
        private bool _isUpdatingAvailableProperties = false; // Flag to prevent premature closure

        public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties, FlatPatternExporterMainWindow mainWindow)
        {
            InitializeComponent();

            _presetIProperties = presetIProperties;
            _mainWindow = mainWindow; // Сохраняем ссылку на MainWindow

            // Инициализируем AvailableProperties, исключая уже присутствующие колонки
            AvailableProperties =
                new ObservableCollection<PresetIProperty>(_presetIProperties.Where(p =>
                    !_mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == p.InternalName)));
            DataContext = this;
            SelectedProperties = new List<PresetIProperty>();

            // Подписка на событие изменения коллекции AvailableProperties
            AvailableProperties.CollectionChanged += AvailableProperties_CollectionChanged;

            // Подписка на событие изменения коллекции PresetIProperties
            _presetIProperties.CollectionChanged += PresetIProperties_CollectionChanged;

            // ItemsSource теперь установлен через CollectionViewSource в XAML
            // для обеспечения группировки по категориям
        }

        public ObservableCollection<PresetIProperty> AvailableProperties { get; set; }
        public List<PresetIProperty> SelectedProperties { get; }

        // Обработчик события изменения коллекции AvailableProperties
        private void AvailableProperties_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (!_isUpdatingAvailableProperties && AvailableProperties.Count == 0)
            {
                // Закрываем окно без предупреждения
                this.Close();
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
                AvailableProperties.Clear();

                foreach (var property in _presetIProperties)
                {
                    if (!_mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == property.InternalName))
                    {
                        AvailableProperties.Add(property);
                    }
                }
                
                // Принудительно обновляем CollectionViewSource для корректного отображения группировки
                var groupedPropertiesResource = FindResource("GroupedProperties") as CollectionViewSource;
                groupedPropertiesResource?.View?.Refresh();
            }
            finally
            {
                _isUpdatingAvailableProperties = false;
            }
        }

        private void iPropertyListBox_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && IPropertyListBox.SelectedItem != null)
            {
                var listBox = sender as ListBox;
                var selectedItem = IPropertyListBox.SelectedItem as PresetIProperty;

                if (listBox != null && selectedItem != null)
                {
                    // Убедимся, что мышь находится над элементом списка, а не над заголовком группы
                    var element = e.OriginalSource as DependencyObject;
                    while (element != null && !(element is ListBoxItem))
                    {
                        element = VisualTreeHelper.GetParent(element);
                    }

                    if (element is ListBoxItem)
                    {
                        DragDrop.DoDragDrop(listBox, selectedItem, DragDropEffects.Move);
                    }
                }
            }
        }

        private void iPropertyListBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(PresetIProperty)) && _isMouseOverDataGrid)
            {
                var droppedProperty = e.Data.GetData(typeof(PresetIProperty)) as PresetIProperty;
                if (droppedProperty != null && !_mainWindow.PartsDataGrid.Columns.Any(c => c.Header.ToString() == droppedProperty.InternalName))
                {
                    // Обновляем список доступных свойств
                    AvailableProperties.Remove(droppedProperty);

                    // Добавляем колонку в DataGrid сразу после добавления свойства
                    _mainWindow.AddIPropertyColumn(droppedProperty);

                    // Добавляем в список выбранных свойств
                    SelectedProperties.Add(droppedProperty);
                }
            }
        }

        private void iPropertyListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (IPropertyListBox.SelectedItem is PresetIProperty selectedProperty)
            {
                if (!SelectedProperties.Contains(selectedProperty))
                {
                    // Добавляем выбранное свойство в список выбранных
                    SelectedProperties.Add(selectedProperty);
                    // Убираем его из доступных свойств
                    AvailableProperties.Remove(selectedProperty);

                    // Обновляем отображение в DataGrid
                    var mainWindow = Application.Current.Windows.OfType<FlatPatternExporterMainWindow>().FirstOrDefault();
                    mainWindow?.AddIPropertyColumn(selectedProperty);
                }
            }
        }

        public void partsDataGrid_DragOver(object sender, DragEventArgs e)
        {
            _isMouseOverDataGrid = true;
        }

        public void partsDataGrid_DragLeave(object sender, DragEventArgs e)
        {
            _isMouseOverDataGrid = false;
        }

        private void CloseWindow_Click(object sender, RoutedEventArgs e)
        {
            // Закрываем окно
            this.Close();
        }

        private void WindowHeader_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void CustomPropertyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Активируем кнопку '+' только если текстовое поле не пустое
            AddCustomPropertyButton.IsEnabled = !string.IsNullOrWhiteSpace(CustomPropertyTextBox.Text);
        }

        private void AddCustomPropertyButton_Click(object sender, RoutedEventArgs e)
        {
            string customPropertyName = CustomPropertyTextBox.Text;
            
            if (!string.IsNullOrWhiteSpace(customPropertyName))
            {
                // Проверяем, не добавлено ли уже это свойство
                if (_mainWindow._customPropertiesList.Contains(customPropertyName))
                {
                    System.Windows.MessageBox.Show($"Свойство '{customPropertyName}' уже добавлено.", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_mainWindow.PartsDataGrid.Columns.Any(c => c.Header as string == customPropertyName))
                {
                    System.Windows.MessageBox.Show($"Столбец с именем '{customPropertyName}' уже существует.", "Предупреждение",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Добавляем в список custom properties
                _mainWindow._customPropertiesList.Add(customPropertyName);

                // Отключаем интерфейс Inventor на время операции
                _mainWindow.SetInventorUserInterfaceState(true);

                try
                {
                    // Создаем колонку (данные заполнятся автоматически)
                    _mainWindow.AddCustomIPropertyColumn(customPropertyName);
                }
                finally
                {
                    _mainWindow.SetInventorUserInterfaceState(false);
                }

                // Очищаем текстовое поле
                CustomPropertyTextBox.Text = "";
            }
        }
    }
}