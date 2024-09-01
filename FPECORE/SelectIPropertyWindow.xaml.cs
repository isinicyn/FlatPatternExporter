using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;

namespace FPECORE
{
    public partial class SelectIPropertyWindow : Window
    {
        public ObservableCollection<PresetIProperty> AvailableProperties { get; set; }
        public List<PresetIProperty> SelectedProperties { get; private set; }
        private ObservableCollection<PresetIProperty> _presetIProperties;
        private bool _isMouseOverDataGrid = false; // Добавили переменную

        public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties)
        {
            InitializeComponent();
            _presetIProperties = presetIProperties;
            AvailableProperties = new ObservableCollection<PresetIProperty>(_presetIProperties.Where(p => !p.IsAdded));
            DataContext = this;
            SelectedProperties = new List<PresetIProperty>();

            // Подписка на событие изменения коллекции PresetIProperties
            _presetIProperties.CollectionChanged += PresetIProperties_CollectionChanged;
        }

        // Обработчик события изменения коллекции
        private void PresetIProperties_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // Обновляем список AvailableProperties при изменении коллекции PresetIProperties
            UpdateAvailableProperties();
        }

        // Обновление списка доступных свойств
        public void UpdateAvailableProperties()
        {
            AvailableProperties.Clear();
            foreach (var property in _presetIProperties.Where(p => !p.IsAdded))
            {
                AvailableProperties.Add(property);
            }
        }

        private void iPropertyListBox_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && iPropertyListBox.SelectedItem != null)
            {
                var listBox = sender as System.Windows.Controls.ListBox;
                if (listBox?.SelectedItem is PresetIProperty selectedProperty)
                {
                    DragDrop.DoDragDrop(listBox, selectedProperty, System.Windows.DragDropEffects.Move);
                }
            }
        }

        private void iPropertyListBox_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(PresetIProperty)) && _isMouseOverDataGrid)
            {
                var droppedProperty = e.Data.GetData(typeof(PresetIProperty)) as PresetIProperty;
                if (droppedProperty != null && !SelectedProperties.Contains(droppedProperty))
                {
                    // Устанавливаем флаг IsAdded у перемещенного свойства и обновляем списки
                    droppedProperty.IsAdded = true;

                    // Обновляем список доступных свойств
                    AvailableProperties.Remove(droppedProperty);

                    // Добавляем колонку в DataGrid сразу после добавления свойства
                    var mainWindow = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                    mainWindow?.AddIPropertyColumn(droppedProperty);

                    // В этом месте можно добавить колонку в partsDataGrid и сразу же обновить список
                    SelectedProperties.Add(droppedProperty);
                }
            }
        }
        private void iPropertyListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (iPropertyListBox.SelectedItem is PresetIProperty selectedProperty)
            {
                if (!SelectedProperties.Contains(selectedProperty))
                {
                    // Добавляем выбранное свойство в список выбранных
                    SelectedProperties.Add(selectedProperty);
                    // Убираем его из доступных свойств
                    AvailableProperties.Remove(selectedProperty);

                    // Обновляем отображение в DataGrid
                    var mainWindow = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                    mainWindow?.AddIPropertyColumn(selectedProperty);
                }
            }
        }

        public void partsDataGrid_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            _isMouseOverDataGrid = true;
        }

        public void partsDataGrid_DragLeave(object sender, System.Windows.DragEventArgs e)
        {
            _isMouseOverDataGrid = false;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}