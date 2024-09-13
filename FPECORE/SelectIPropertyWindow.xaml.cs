using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Input;
using Application = System.Windows.Application;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using ListBox = System.Windows.Controls.ListBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;

namespace FPECORE;

public partial class SelectIPropertyWindow : Window
{
    private bool _isMouseOverDataGrid; // Добавили переменную
    private readonly MainWindow _mainWindow; // Поле для хранения ссылки на MainWindow
    private readonly ObservableCollection<PresetIProperty> _presetIProperties;

    public SelectIPropertyWindow(ObservableCollection<PresetIProperty> presetIProperties, MainWindow mainWindow)
    {
        InitializeComponent();
        _presetIProperties = presetIProperties;
        _mainWindow = mainWindow; // Сохраняем ссылку на MainWindow
        AvailableProperties =
            new ObservableCollection<PresetIProperty>(_presetIProperties.Where(p =>
                !_mainWindow.IsColumnPresent(p.InternalName)));
        DataContext = this;
        SelectedProperties = new List<PresetIProperty>();

        // Подписка на событие изменения коллекции PresetIProperties
        _presetIProperties.CollectionChanged += PresetIProperties_CollectionChanged;
    }

    public ObservableCollection<PresetIProperty> AvailableProperties { get; set; }
    public List<PresetIProperty> SelectedProperties { get; }

    // Обработчик события изменения коллекции
    private void PresetIProperties_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
    {
        // Обновляем список AvailableProperties при изменении коллекции PresetIProperties
        UpdateAvailableProperties();
    }

    // Обновление списка доступных свойств
    public void UpdateAvailableProperties()
    {
        AvailableProperties.Clear();

        foreach (var property in _presetIProperties)
            // Используем метод IsColumnPresent через ссылку на MainWindow
            if (!_mainWindow.IsColumnPresent(property.InternalName))
                AvailableProperties.Add(property);
    }

    private void iPropertyListBox_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed && iPropertyListBox.SelectedItem != null)
        {
            var listBox = sender as ListBox;
            if (listBox?.SelectedItem is PresetIProperty selectedProperty)
                DragDrop.DoDragDrop(listBox, selectedProperty, DragDropEffects.Move);
        }
    }

    private void iPropertyListBox_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(typeof(PresetIProperty)) && _isMouseOverDataGrid)
        {
            var droppedProperty = e.Data.GetData(typeof(PresetIProperty)) as PresetIProperty;
            if (droppedProperty != null && !_mainWindow.IsColumnPresent(droppedProperty.InternalName))
            {
                // Обновляем список доступных свойств
                AvailableProperties.Remove(droppedProperty);

                // Добавляем колонку в DataGrid сразу после добавления свойства
                _mainWindow.AddIPropertyColumn(droppedProperty);

                // В этом месте можно добавить колонку в partsDataGrid и сразу же обновить список
                SelectedProperties.Add(droppedProperty);
            }
        }
    }

    private void iPropertyListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (iPropertyListBox.SelectedItem is PresetIProperty selectedProperty)
            if (!SelectedProperties.Contains(selectedProperty))
            {
                // Добавляем выбранное свойство в список выбранных
                SelectedProperties.Add(selectedProperty);
                // Убираем его из доступных свойств
                AvailableProperties.Remove(selectedProperty);

                // Обновляем отображение в DataGrid
                var mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                mainWindow?.AddIPropertyColumn(selectedProperty);
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

}