using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using DefineEdge;
using Inventor;
using LayerSettingsApp;
using Application = Inventor.Application;
using Binding = System.Windows.Data.Binding;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using Color = System.Windows.Media.Color;
using DragEventArgs = System.Windows.DragEventArgs;
using File = System.IO.File;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Image = System.Windows.Controls.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Path = System.IO.Path;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;
using TreeView = System.Windows.Controls.TreeView;

namespace FPECORE;

public partial class MainWindow : Window
{
    private const string PlaceholderText = "Поиск...";
    private AdornerLayer _adornerLayer;
    private HeaderAdorner _headerAdorner;
    private bool _isColumnDraggedOutside;
    private DataGridColumn _reorderingColumn;
    private string _actualSearchText = string.Empty; // Поле для хранения фактического текста поиска

    private readonly List<string> _customPropertiesList = new();
    private string _fixedFolderPath = string.Empty;
    private bool _isCancelled;

    private bool _isCtrlPressed;
    private bool _isEditing;
    private bool _isScanning;
    private int _itemCounter = 1; // Инициализация счетчика пунктов
    private Document? _lastScannedDocument;
    private List<PartData> _originalPartsData = new(); // Для хранения исходного состояния
    private readonly ObservableCollection<PartData> _partsData = new();
    private readonly CollectionViewSource _partsDataView;
    private readonly DispatcherTimer _searchDelayTimer;
    private Application _thisApplication;
    private List<PartData> _conflictingParts = new(); // Список конфликтующих файлов
    private Dictionary<string, List<(string PartNumber, string FileName)>> _conflictFileDetails = new();

    public MainWindow()
    {
        InitializeComponent();
        InitializeInventor();
        InitializeProjectFolder(); // Инициализация папки проекта при запуске
        PartsDataGrid.ItemsSource = _partsData;
        MultiplierTextBox.IsEnabled = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        UpdateFileCountLabel(0);
        ClearButton.IsEnabled = false;
        PartsDataGrid.PreviewMouseDown += PartsDataGrid_PreviewMouseDown;
        PartsDataGrid.PreviewMouseMove += PartsDataGrid_PreviewMouseMove;
        PartsDataGrid.PreviewMouseLeftButtonUp += PartsDataGrid_PreviewMouseLeftButtonUp;
        PartsDataGrid.ColumnReordering += PartsDataGrid_ColumnReordering;
        PartsDataGrid.ColumnReordered += PartsDataGrid_ColumnReordered;
        ChooseFolderRadioButton.IsChecked = true;

        PartsDataGrid.DragOver += PartsDataGrid_DragOver;
        PartsDataGrid.DragLeave += PartsDataGrid_DragLeave;

        // Инициализация CollectionViewSource для фильтрации
        _partsDataView = new CollectionViewSource { Source = _partsData };
        _partsDataView.Filter += PartsData_Filter;

        // Устанавливаем ItemsSource для DataGrid
        PartsDataGrid.ItemsSource = _partsDataView.View;

        // Подписываемся на событие изменения коллекции
        _partsData.CollectionChanged += PartsData_CollectionChanged;

        UpdateEditButtonState(); // Обновляем состояние кнопки после загрузки данных

        // Настраиваем таймер для задержки поиска
        _searchDelayTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1) // 1 секунда задержки
        };
        _searchDelayTimer.Tick += SearchDelayTimer_Tick;

        // Установка начального текста для searchTextBox
        SearchTextBox.Text = PlaceholderText;
        SearchTextBox.Foreground = Brushes.Gray;

        // Создаем экземпляр MainWindow из LayerSettingsApp
        var layerSettingsWindow = new LayerSettingsApp.MainWindow();

        // Используем его данные для настройки LayerSettings
        LayerSettings = layerSettingsWindow.LayerSettings;
        AvailableColors = layerSettingsWindow.AvailableColors;
        LineTypes = layerSettingsWindow.LineTypes;


        // Инициализация предустановленных колонок
        PresetIProperties = new ObservableCollection<PresetIProperty>
        {
            new() { InternalName = "ID", DisplayName = "#Нумерация", InventorPropertyName = "Item" },
            new() { InternalName = "Обозначение", DisplayName = "Обозначение", InventorPropertyName = "PartNumber" },
            new() { InternalName = "Наименование", DisplayName = "Наименование", InventorPropertyName = "Description" },
            new() { InternalName = "Состояние модели", DisplayName = "Состояние модели", InventorPropertyName = "ModelState" },
            new() { InternalName = "Материал", DisplayName = "Материал", InventorPropertyName = "Material" },
            new() { InternalName = "Толщина", DisplayName = "Толщина", InventorPropertyName = "Thickness" },
            new() { InternalName = "Количество", DisplayName = "Количество", InventorPropertyName = "Quantity" },
            new() { InternalName = "Изображение детали", DisplayName = "Изображение детали", InventorPropertyName = "Preview" },
            new() { InternalName = "Изображение развертки", DisplayName = "Изображение развертки", InventorPropertyName = "DxfPreview" },

            // Добавление новых предустановленных свойств
            new() { InternalName = "Автор", DisplayName = "Автор", InventorPropertyName = "Author" }, //, ShouldBeAddedOnInit = true or false
            new() { InternalName = "Ревизия", DisplayName = "Ревизия", InventorPropertyName = "Revision Number" },
            new() { InternalName = "Проект", DisplayName = "Проект", InventorPropertyName = "Project" },
            new() { InternalName = "Инвентарный номер", DisplayName = "Инвентарный номер", InventorPropertyName = "Stock Number" }
        };

        // Устанавливаем DataContext для текущего окна, объединяя данные из LayerSettingsWindow и других источников
        DataContext = this;

        // Инициализация DataGrid с предустановленными колонками
        InitializePresetColumns();

        // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
        KeyDown += MainWindow_KeyDown;
        KeyUp += MainWindow_KeyUp;

        VersionTextBlock.Text = "Версия программы: " + GetVersion();
    }

    public ObservableCollection<LayerSetting> LayerSettings { get; set; }
    public ObservableCollection<string> AvailableColors { get; set; }
    public ObservableCollection<string> LineTypes { get; set; }
    public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }

    public string GetVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();

        // Попытка получить информацию о версии из различных атрибутов
        var informationalVersion =
            assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
        var assemblyVersion = assembly.GetName().Version?.ToString();

        // Если версия содержит commit hash, обрезаем его до 7 символов
        if (informationalVersion != null && informationalVersion.Contains("+"))
        {
            // Разбиваем строку на части и обрезаем commit hash до 7 символов
            var parts = informationalVersion.Split('+');
            var shortCommitHash = parts[1].Length > 7 ? parts[1].Substring(0, 7) : parts[1];
            informationalVersion = $"{parts[0]}+{shortCommitHash}";
        }

        // Возвращаем обрезанную версию или другие версии
        return informationalVersion ?? fileVersion ?? assemblyVersion ?? "Версия неизвестна";
    }

    public bool IsColumnPresent(string columnName)
    {
        var result = false;

        // Используем Dispatcher для доступа к UI-элементам
        Dispatcher.Invoke(() => { result = PartsDataGrid.Columns.Any(c => c.Header.ToString() == columnName); });

        return result;
    }

    private void InitializePresetColumns()
    {
        if (PresetIProperties == null) throw new InvalidOperationException("PresetIProperties is not initialized.");

        foreach (var property in PresetIProperties)
        {
            // Если колонка уже есть в таблице, ничего не делаем
            if (IsColumnPresent(property.InternalName)) continue;

            // Если колонка не добавлена в XAML, пропускаем её добавление
            if (property.ShouldBeAddedOnInit) AddIPropertyColumn(property);
        }
    }

    private async Task<string> GetPropertyExpressionOrValueAsync(PartDocument partDoc, string propertyName,
        bool getExpression = false)
    {
        return await Task.Run(() =>
        {
            try
            {
                var propertySets = partDoc.PropertySets;
                foreach (PropertySet propSet in propertySets)
                foreach (Property prop in propSet)
                    if (prop.Name == propertyName)
                    {
                        if (getExpression)
                            // Попытка получить выражение
                            try
                            {
                                var expression = prop.Expression;
                                if (!string.IsNullOrEmpty(expression)) return expression; // Возвращаем выражение
                            }
                            catch (Exception)
                            {
                                // Если нет выражения, вернем значение
                                return prop.Value?.ToString() ?? string.Empty;
                            }

                        return prop.Value?.ToString() ?? string.Empty;
                    }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при получении выражения или значения для {propertyName}: {ex.Message}");
            }

            return string.Empty;
        });
    }

    private async void EditButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isEditing)
        {
            // Отключаем режим редактирования
            _isEditing = false;
            PartsDataGrid.IsReadOnly = true;
            EditButton.Content = "Редактировать";
            SaveButton.IsEnabled = false;
            ExportButton.IsEnabled = true; // Разблокируем кнопку экспорта

            // Восстанавливаем исходное состояние данных
            for (var i = 0; i < _partsData.Count; i++)
            {
                var currentItem = _partsData[i];
                var originalItem = _originalPartsData[i];

                currentItem.PartNumber = originalItem.PartNumber; // Восстанавливаем значение
                currentItem.OriginalPartNumber =
                    originalItem.OriginalPartNumber; // Восстанавливаем оригинальный PartNumber
                currentItem.Description = originalItem.Description; // Восстанавливаем значение

                // Восстанавливаем кастомные свойства
                foreach (var key in currentItem.CustomProperties.Keys.ToList())
                    currentItem.CustomProperties[key] = originalItem.CustomProperties.ContainsKey(key)
                        ? originalItem.CustomProperties[key]
                        : currentItem.CustomProperties[key];

                // Восстанавливаем флаги
                currentItem.IsModified = originalItem.IsModified;
                currentItem.IsReadOnly = originalItem.IsReadOnly;
            }

            PartsDataGrid.Items.Refresh(); // Обновляем таблицу
        }
        else
        {
            // Включаем режим редактирования
            _isEditing = true;
            PartsDataGrid.IsReadOnly = false;
            EditButton.Content = "Отмена";
            SaveButton.IsEnabled = true;
            ExportButton.IsEnabled = false; // Блокируем кнопку экспорта в режиме редактирования

            // Сохраняем исходное состояние данных
            _originalPartsData = _partsData.Select(item => new PartData
            {
                PartNumber = item.PartNumber,
                OriginalPartNumber = item.OriginalPartNumber, // Сохраняем оригинальный PartNumber
                PartNumberExpression = item.PartNumberExpression,
                Description = item.Description,
                DescriptionExpression = item.DescriptionExpression,
                CustomProperties = new Dictionary<string, string>(item.CustomProperties),
                CustomPropertyExpressions = new Dictionary<string, string>(item.CustomPropertyExpressions),
                IsModified = item.IsModified, // Сохраняем состояние флага
                IsReadOnly = item.IsReadOnly // Сохраняем состояние флага
            }).ToList();

            // Разрешаем редактирование и сбрасываем флаг изменений
            foreach (var item in _partsData)
            {
                item.IsModified = false; // Обнуляем флаг изменений
                item.OriginalPartNumber = item.PartNumber; // Сохраняем оригинальный PartNumber перед редактированием
                item.IsReadOnly = false; // Разрешаем редактирование
            }

            // Если чекбокс активен, отображаем выражения
            if (ExpressionsCheckBox.IsChecked == true)
                foreach (var item in _partsData)
                {
                    item.PartNumber = item.PartNumberExpression; // Показываем выражение
                    item.Description = item.DescriptionExpression; // Показываем выражение

                    foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        // Проверяем наличие ключа перед доступом к выражению
                        if (item.CustomPropertyExpressions.ContainsKey(customProperty))
                            item.CustomProperties[customProperty] =
                                item.CustomPropertyExpressions[customProperty]; // Показываем выражение
                        else
                            item.CustomProperties[customProperty] =
                                item.CustomProperties[customProperty]; // Оставляем значение
                }

            PartsDataGrid.Items.Refresh(); // Обновляем таблицу
        }

        UpdateEditButtonState();
    }

    private void ExpressionsCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        if (_isEditing) // Применяем изменения только в режиме редактирования
        {
            if (ExpressionsCheckBox.IsChecked == true)
            {
                // Показываем выражения
                foreach (var item in _partsData)
                {
                    item.SetPartNumberWithoutModification(item.PartNumberExpression); // Показываем выражение
                    item.SetDescriptionWithoutModification(item.DescriptionExpression); // Показываем выражение

                    foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        item.SetCustomPropertyWithoutModification(customProperty, item.CustomPropertyExpressions[customProperty]); // Показываем выражение
                }
            }
            else
            {
                // Показываем обычные значения
                foreach (var item in _partsData)
                {
                    item.SetPartNumberWithoutModification(item.PartNumber); // Показываем значение
                    item.SetDescriptionWithoutModification(item.Description); // Показываем значение

                    foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        item.SetCustomPropertyWithoutModification(customProperty, item.CustomProperties[customProperty]); // Показываем значение
                }
            }

            PartsDataGrid.Items.Refresh(); // Обновляем таблицу после изменения
        }
    }

    private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        _isEditing = false;
        PartsDataGrid.IsReadOnly = true;
        EditButton.Content = "Редактировать";
        SaveButton.IsEnabled = false;

        foreach (var item in _partsData)
        {
            if (item.IsModified) // Сохраняем только изменённые элементы
                await SaveIPropertyChangesAsync(item);

            // После успешного сохранения сбрасываем флаг IsModified и делаем строки только для чтения
            item.IsModified = false;
            item.IsReadOnly = true;
        }

        _originalPartsData.Clear(); // Очищаем сохранённые оригинальные данные
        PartsDataGrid.Items.Refresh();

        // Восстанавливаем состояние кнопки
        UpdateEditButtonState();
    }

    private async Task SaveIPropertyChangesAsync(PartData item)
    {
        try
        {
            var partDoc = OpenPartDocument(item.OriginalPartNumber);
            if (partDoc != null)
            {
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number", item.PartNumber);
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description", item.Description);

                await Task.Run(() =>
                {
                    partDoc.Save2();
                    partDoc.Close();
                });

                // После успешного сохранения сбрасываем флаг
                item.IsModified = false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void PartsDataGrid_DragOver(object sender, DragEventArgs e)
    {
        // Уведомляем SelectIPropertyWindow, что курсор находится над DataGrid
        var selectIPropertyWindow =
            System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
        if (selectIPropertyWindow != null) selectIPropertyWindow.partsDataGrid_DragOver(sender, e);
    }

    private void PartsDataGrid_DragLeave(object sender, DragEventArgs e)
    {
        // Уведомляем SelectIPropertyWindow, что курсор ушел из DataGrid
        var selectIPropertyWindow =
            System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
        if (selectIPropertyWindow != null) selectIPropertyWindow.partsDataGrid_DragLeave(sender, e);
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);

        // Закрытие всех дочерних окон
        foreach (Window window in System.Windows.Application.Current.Windows)
            if (window != this)
                window.Close();

        // Если были созданы дополнительные потоки, убедитесь, что они завершены
    }

    private void PartsData_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
    {
        UpdateEditButtonState(); // Обновляем состояние кнопки при изменении данных
    }

    private void UpdateEditButtonState()
    {
        // Если данные в таблице присутствуют, и не активирован режим редактирования
        if (_partsData != null && _partsData.Count > 0)
        {
            EditButton.IsEnabled = true; // Кнопка "Редактировать/Отмена" всегда активна, если есть данные
            SaveButton.IsEnabled = _isEditing; // Кнопка "Сохранить" активна только в режиме редактирования
        }
        else
        {
            EditButton.IsEnabled = false; // Деактивируем кнопку, если данных нет
            SaveButton.IsEnabled = false; // Деактивируем кнопку сохранения, если данных нет
        }
    }

    private void AddPresetIPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        var selectIPropertyWindow =
            new SelectIPropertyWindow(PresetIProperties, this); // Передаем `this` как ссылку на MainWindow
        selectIPropertyWindow.ShowDialog();

        // После закрытия окна добавляем выбранные свойства
        foreach (var property in selectIPropertyWindow.SelectedProperties) AddIPropertyColumn(property);
    }

    private void MainWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        // Получаем исходный элемент
        var clickedElement = e.OriginalSource as DependencyObject;

        // Поднимаемся по дереву визуальных элементов
        while (clickedElement != null && !(clickedElement is Visual || clickedElement is Visual3D))
        {
            // Проверяем, что это не элемент типа Run
            if (clickedElement is Run)
                // Просто выходим, если наткнулись на Run
                return;
            clickedElement = VisualTreeHelper.GetParent(clickedElement);
        }

        // Получаем источник события
        var source = e.OriginalSource as DependencyObject;

        // Ищем DataGrid по иерархии визуальных элементов
        while (source != null && !(source is DataGrid)) source = VisualTreeHelper.GetParent(source);

        // Если DataGrid не найден (клик был вне DataGrid), сбрасываем выделение
        if (source == null) PartsDataGrid.UnselectAll();
    }

    private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
        // Очищаем текстовое поле и восстанавливаем PlaceholderText
        SearchTextBox.Text = string.Empty;
        SearchTextBox.Text = PlaceholderText;
        SearchTextBox.Foreground = Brushes.Gray;
        ClearSearchButton.Visibility = Visibility.Collapsed;
        SearchTextBox.CaretIndex = 0; // Перемещаем курсор в начало
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (SearchTextBox.Text == PlaceholderText)
        {
            _actualSearchText = string.Empty;
            ClearSearchButton.Visibility = Visibility.Collapsed;
            return;
        }

        _actualSearchText = SearchTextBox.Text.Trim().ToLower();

        // Показываем или скрываем кнопку очистки в зависимости от наличия текста
        ClearSearchButton.Visibility =
            string.IsNullOrEmpty(_actualSearchText) ? Visibility.Collapsed : Visibility.Visible;

        // Перезапуск таймера при изменении текста
        _searchDelayTimer.Stop();
        _searchDelayTimer.Start();
    }

    private void SearchTextBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (SearchTextBox.Text == PlaceholderText)
        {
            SearchTextBox.Text = string.Empty;
            SearchTextBox.Foreground = Brushes.Black;
        }
    }

    private void SearchTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(SearchTextBox.Text))
        {
            SearchTextBox.Text = PlaceholderText;
            SearchTextBox.Foreground = Brushes.Gray;
            _actualSearchText = string.Empty; // Очищаем фактический текст поиска
            ClearSearchButton.Visibility = Visibility.Collapsed; // Скрываем кнопку очистки
        }
    }

    private void SearchDelayTimer_Tick(object sender, EventArgs e)
    {
        _searchDelayTimer.Stop();
        _partsDataView.View.Refresh(); // Обновление фильтрации
    }

    private void PartsData_Filter(object sender, FilterEventArgs e)
    {
        if (e.Item is PartData partData)
        {
            if (string.IsNullOrEmpty(_actualSearchText))
            {
                e.Accepted = true;
                return;
            }

            var matches = partData.PartNumber?.ToLower().Contains(_actualSearchText) == true ||
                          partData.Description?.ToLower().Contains(_actualSearchText) == true ||
                          partData.ModelState?.ToLower().Contains(_actualSearchText) == true ||
                          partData.Material?.ToLower().Contains(_actualSearchText) == true ||
                          partData.Thickness?.ToLower().Contains(_actualSearchText) == true ||
                          partData.Quantity.ToString().Contains(_actualSearchText);

            // Проверка на пользовательские свойства
            if (!matches)
                foreach (var property in partData.CustomProperties)
                    if (property.Value.ToLower().Contains(_actualSearchText))
                    {
                        matches = true;
                        break;
                    }

            e.Accepted = matches;
        }
        else
        {
            e.Accepted = false;
        }
    }

    private void partsDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
    {
        // Устанавливаем чередование цветов
        if (e.Row.GetIndex() % 2 == 0)
            e.Row.Background = new SolidColorBrush(Colors.White);
        else
            e.Row.Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)); // #F5F5F5
    }

    private void partsDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(typeof(PresetIProperty)))
        {
            var droppedData = e.Data.GetData(typeof(PresetIProperty)) as PresetIProperty;
            if (droppedData != null && !IsColumnPresent(droppedData.InternalName))
            {
                // Определяем позицию мыши
                var position = e.GetPosition(PartsDataGrid);

                // Найти колонку, перед которой должна быть вставлена новая колонка
                var insertIndex = -1;
                double totalWidth = 0;

                for (var i = 0; i < PartsDataGrid.Columns.Count; i++)
                {
                    totalWidth += PartsDataGrid.Columns[i].ActualWidth;
                    if (position.X < totalWidth)
                    {
                        insertIndex = i;
                        break;
                    }
                }

                // Добавляем колонку на нужную позицию
                AddIPropertyColumn(droppedData, insertIndex);
            }
        }
    }

    public void AddIPropertyColumn(PresetIProperty iProperty, int? insertIndex = null)
    {
        // Проверяем, существует ли уже колонка с таким заголовком
        if (IsColumnPresent(iProperty.InternalName))
            return;

        DataGridColumn column;

        // Проверка, если это колонка с изображением детали или развертки
        if (iProperty.InternalName == "Изображение детали" || iProperty.InternalName == "Изображение развертки")
        {
            column = new DataGridTemplateColumn
            {
                Header = iProperty.InternalName,
                CellTemplate = new DataTemplate
                {
                    VisualTree = new FrameworkElementFactory(typeof(Image))
                }
            };
            // Устанавливаем привязку данных для колонки изображения
            (column as DataGridTemplateColumn).CellTemplate.VisualTree.SetBinding(Image.SourceProperty,
                new Binding(iProperty.InventorPropertyName));
        }
        else
        {
            // Создаем текстовую колонку
            var textColumn = new DataGridTextColumn
            {
                Header = iProperty.InternalName,
                Binding = new Binding(iProperty.InventorPropertyName)
            };

            // Применяем специальный стиль для колонки "Количество"
            if (iProperty.InternalName == "Количество")
                textColumn.ElementStyle = PartsDataGrid.FindResource("QuantityCellStyle") as Style;
            else
                // Применяем стиль CenteredCellStyle для всех остальных текстовых колонок
                textColumn.ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style;

            column = textColumn;
        }

        // Если указан индекс вставки, вставляем колонку в нужное место
        if (insertIndex.HasValue && insertIndex.Value >= 0 && insertIndex.Value < PartsDataGrid.Columns.Count)
            PartsDataGrid.Columns.Insert(insertIndex.Value, column);
        else
            // В противном случае добавляем колонку в конец
            PartsDataGrid.Columns.Add(column);

        // Обновляем список доступных свойств
        var selectIPropertyWindow =
            System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
        selectIPropertyWindow?.UpdateAvailableProperties();
    }

    private void AddPartData(PartData partData)
    {
        _partsData.Add(partData);
        _partsDataView.View.Refresh(); // Обновляем фильтрацию и отображение

        // Обновляем состояние кнопки после добавления данных
        UpdateEditButtonState();
    }

    private void PartsDataGrid_ColumnReordering(object sender, DataGridColumnReorderingEventArgs e)
    {
        _reorderingColumn = e.Column;
        _isColumnDraggedOutside = false;

        // Создаем фантомный заголовок с фиксированным размером
        var header = new TextBlock
        {
            Text = _reorderingColumn.Header.ToString(),
            Background = Brushes.LightGray,
            FontSize = 12, // Установите желаемый размер шрифта
            Padding = new Thickness(5),
            Opacity = 0.7, // Полупрозрачность
            TextAlignment = TextAlignment.Center,
            Width = _reorderingColumn.ActualWidth,
            Height = 30,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            IsHitTestVisible = false
        };

        // Создаем и добавляем Adorner
        _adornerLayer = AdornerLayer.GetAdornerLayer(PartsDataGrid);
        _headerAdorner = new HeaderAdorner(PartsDataGrid, header);
        _adornerLayer.Add(_headerAdorner);

        // Изначально размещаем Adorner там, где находится заголовок
        UpdateAdornerPosition(e);
    }

    private void PartsDataGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed)
        {
            // Определяем, кликнул ли пользователь на заголовке колонки
            var header = e.OriginalSource as FrameworkElement;
            if (header?.Parent is DataGridColumnHeader columnHeader)
                // Начинаем перетаскивание
                StartColumnReordering(columnHeader.Column);
        }
    }

    private void StartColumnReordering(DataGridColumn column)
    {
        _reorderingColumn = column;
        _isColumnDraggedOutside = false;

        // Создаем фантомный заголовок с фиксированным размером
        var header = new TextBlock
        {
            Text = _reorderingColumn.Header.ToString(),
            Background = Brushes.LightGray,
            FontSize = 12, // Установите желаемый размер шрифта
            Padding = new Thickness(5),
            Opacity = 0.7, // Полупрозрачность
            TextAlignment = TextAlignment.Center,
            Width = _reorderingColumn.ActualWidth,
            Height = 30,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            IsHitTestVisible = false
        };

        // Создаем и добавляем Adorner
        _adornerLayer = AdornerLayer.GetAdornerLayer(PartsDataGrid);
        _headerAdorner = new HeaderAdorner(PartsDataGrid, header);
        _adornerLayer.Add(_headerAdorner);

        // Изначально размещаем Adorner там, где находится заголовок
        var position = Mouse.GetPosition(PartsDataGrid);
        _headerAdorner.RenderTransform = new TranslateTransform(
            position.X - _headerAdorner.DesiredSize.Width / 2,
            position.Y - _headerAdorner.DesiredSize.Height / 2
        );
    }

    private void PartsDataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (_headerAdorner == null || _reorderingColumn == null)
            return;

        // Обновляем позицию Adorner
        UpdateAdornerPosition(null);

        // Получаем текущую позицию мыши относительно DataGrid
        var position = e.GetPosition(PartsDataGrid);

        // Проверяем, если мышь вышла за границы DataGrid
        if (position.X > PartsDataGrid.ActualWidth || position.Y < 0 || position.Y > PartsDataGrid.ActualHeight)
        {
            _isColumnDraggedOutside = true;
            SetAdornerColor(Brushes.Red); // Меняем цвет на красный и добавляем надпись "УДАЛИТЬ"
        }
        else
        {
            _isColumnDraggedOutside = false;
            SetAdornerColor(Brushes.LightGray); // Возвращаем исходный цвет и удаляем надпись "УДАЛИТЬ"
        }
    }

    private void SetAdornerColor(Brush color)
    {
        if (_headerAdorner != null && _headerAdorner.Child is TextBlock textBlock)
        {
            textBlock.Background = color;

            if (color == Brushes.Red)
            {
                textBlock.Text = $"❎ {_reorderingColumn.Header}";
                textBlock.Foreground = Brushes.White; // Белый текст на красном фоне
                textBlock.FontWeight = FontWeights.Bold; // Жирный шрифт для улучшенной видимости
                textBlock.Padding = new Thickness(5); // Дополнительное внутреннее пространство
                textBlock.TextAlignment = TextAlignment.Center; // Центрирование текста
                textBlock.VerticalAlignment = VerticalAlignment.Center;
                textBlock.HorizontalAlignment = HorizontalAlignment.Center;
            }
            else
            {
                textBlock.Text = _reorderingColumn.Header.ToString();
                textBlock.Foreground = Brushes.Black; // Черный текст на сером фоне
                textBlock.FontWeight = FontWeights.Normal;
                textBlock.Padding = new Thickness(5);
            }

            // Убираем фиксированные размеры и устанавливаем авторазмеры
            textBlock.Width = double.NaN;
            textBlock.Height = double.NaN;
            textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            textBlock.Arrange(new Rect(textBlock.DesiredSize));
        }
    }

    private void PartsDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_adornerLayer != null && _headerAdorner != null)
        {
            _adornerLayer.Remove(_headerAdorner);
            _headerAdorner = null;
        }

        if (_reorderingColumn != null && _isColumnDraggedOutside)
        {
            var columnName = _reorderingColumn.Header.ToString();

            // Убираем колонку из DataGrid
            var columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == columnName);
            if (columnToRemove != null)
            {
                PartsDataGrid.Columns.Remove(columnToRemove);

                // Удаляем кастомное свойство из списка и данных деталей
                RemoveCustomIPropertyColumn(columnName);

                // Обновляем список доступных свойств, если это кастомное свойство
                var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>()
                    .FirstOrDefault();
                selectIPropertyWindow?.UpdateAvailableProperties();
            }
        }

        _reorderingColumn = null;
        _isColumnDraggedOutside = false;
    }

    private void PartsDataGrid_ColumnReordered(object sender, DataGridColumnEventArgs e)
    {
        _reorderingColumn = null;
        _isColumnDraggedOutside = false;

        // Убираем Adorner
        if (_adornerLayer != null && _headerAdorner != null)
        {
            _adornerLayer.Remove(_headerAdorner);
            _headerAdorner = null;
        }
    }

    private void UpdateAdornerPosition(DataGridColumnReorderingEventArgs e)
    {
        if (_headerAdorner == null)
            return;

        // Получаем текущую позицию мыши относительно родителя DataGrid
        var position = Mouse.GetPosition(_adornerLayer);

        // Применяем трансформацию для перемещения Adorner
        _headerAdorner.RenderTransform = new TranslateTransform(
            position.X - _headerAdorner.RenderSize.Width / 2,
            position.Y - _headerAdorner.RenderSize.Height / 1.8
        );
    }

    private void InitializeInventor()
    {
        try
        {
            _thisApplication = (Application)MarshalCore.GetActiveObject("Inventor.Application") ??
                              throw new InvalidOperationException("Не удалось получить активный объект Inventor.");
        }
        catch (COMException)
        {
            MessageBox.Show(
                "Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Произошла ошибка при подключении к Inventor: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
            Close();
        }
    }

    private void MainWindow_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
        {
            _isCtrlPressed = true;
            ExportButton.IsEnabled = true;
        }
    }

    private void MainWindow_KeyUp(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
        {
            _isCtrlPressed = false;
            ExportButton.IsEnabled = _partsData.Count > 0; // Восстанавливаем исходное состояние кнопки
        }
    }

    private double GetColumnStartX(DataGrid dataGrid, int columnIndex)
    {
        double startX = 0;
        for (var i = 0; i < columnIndex; i++) startX += dataGrid.Columns[i].ActualWidth;
        return startX;
    }

    private void EnableSplineReplacementCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        var isChecked = EnableSplineReplacementCheckBox.IsChecked == true;
        SplineReplacementComboBox.IsEnabled = isChecked;
        SplineToleranceTextBox.IsEnabled = isChecked;

        if (isChecked && SplineReplacementComboBox.SelectedIndex == -1)
            SplineReplacementComboBox.SelectedIndex = 0; // По умолчанию выбираем "Линии"
    }

    private void PrepareExportOptions(out string options)
    {
        var sb = new StringBuilder();

        sb.Append($"AcadVersion={((ComboBoxItem)AcadVersionComboBox.SelectedItem).Content}");

        if (EnableSplineReplacementCheckBox.IsChecked == true)
        {
            var splineTolerance = SplineToleranceTextBox.Text;

            // Получаем текущий разделитель дробной части из системных настроек
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];

            // Заменяем разделитель на актуальный
            splineTolerance = splineTolerance.Replace('.', decimalSeparator).Replace(',', decimalSeparator);

            if (SplineReplacementComboBox.SelectedIndex == 0) // Линии
                sb.Append($"&SimplifySplines=True&SplineTolerance={splineTolerance}");
            else if (SplineReplacementComboBox.SelectedIndex == 1) // Дуги
                sb.Append($"&SimplifySplines=True&SimplifyAsTangentArcs=True&SplineTolerance={splineTolerance}");
        }
        else
        {
            sb.Append("&SimplifySplines=False");
        }

        if (MergeProfilesIntoPolylineCheckBox.IsChecked == true) sb.Append("&MergeProfilesIntoPolyline=True");

        if (RebaseGeometryCheckBox.IsChecked == true) sb.Append("&RebaseGeometry=True");

        if (TrimCenterlinesCheckBox.IsChecked == true) // Проверяем новое поле TrimCenterlinesAtContour
            sb.Append("&TrimCenterlinesAtContour=True");

        options = sb.ToString();
    }

    private void MultiplierTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        // Проверяем текущий текст вместе с новым вводом
        var newText = (sender as TextBox).Text.Insert((sender as TextBox).CaretIndex, e.Text);

        // Проверяем, является ли текст положительным числом больше нуля
        e.Handled = !IsTextAllowed(newText);
    }

    private static bool IsTextAllowed(string text)
    {
        // Регулярное выражение для проверки положительных чисел
        var regex = new Regex("^[1-9][0-9]*$"); // Разрешены только положительные числа больше нуля
        return regex.IsMatch(text);
    }

    private void UpdateQuantitiesWithMultiplier(int multiplier)
    {
        foreach (var partData in _partsData)
        {
            partData.IsOverridden = false; // Сброс переопределённых значений
            partData.Quantity = partData.OriginalQuantity * multiplier;
            partData.QuantityColor = multiplier > 1 ? Brushes.Blue : Brushes.Black;
        }

        PartsDataGrid.Items.Refresh();
    }

    private async void ScanButton_Click(object sender, RoutedEventArgs e)
    {
        ResetEditMode(); // Сброс режима редактирования перед сканированием

        if (_thisApplication == null)
        {
            MessageBoxHelper.ShowInventorNotRunningError();
            return;
        }

        // Проверка на наличие активного документа
        Document? doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBoxHelper.ShowNoDocumentOpenWarning();
            return;
        }

        var isBackgroundMode = BackgroundModeCheckBox.IsChecked == true;
        if (isBackgroundMode) _thisApplication.UserInterfaceManager.UserInteractionDisabled = true;

        ModelStateInfoRunBottom.Text = string.Empty;

        if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject &&
            doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
        {
            MessageBox.Show("Откройте сборку или деталь для сканирования.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        ResetProgressBar();
        ProgressBar.IsIndeterminate = true;
        ProgressLabel.Text = "Статус: Сбор данных...";
        ScanButton.IsEnabled = false;
        CancelButton.IsEnabled = true;
        ExportButton.IsEnabled = false;
        ClearButton.IsEnabled = false;
        _isScanning = true;
        _isCancelled = false;

        var stopwatch = Stopwatch.StartNew();

        await Task.Delay(100);

        var partCount = 0;
        var documentType = doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "Сборка" : "Деталь";
        var partNumber = string.Empty;
        var description = string.Empty;
        var modelStateInfo = string.Empty;

        _partsData.Clear();
        _itemCounter = 1;
        UpdateEditButtonState(); // Обновляем состояние кнопки редактирования после очистки данных

        var progress = new Progress<PartData>(partData =>
        {
            partData.Item = _itemCounter;
            _partsData.Add(partData);
            PartsDataGrid.Items.Refresh();
            _itemCounter++;
        });

        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)doc;
            var sheetMetalParts = new Dictionary<string, int>();
            if (TraverseRadioButton.IsChecked == true)
                await Task.Run(() =>
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));
            partCount = sheetMetalParts.Count;
            partNumber = GetProperty(asmDoc.PropertySets["Design Tracking Properties"], "Part Number");
            description = GetProperty(asmDoc.PropertySets["Design Tracking Properties"], "Description");
            modelStateInfo = asmDoc.ComponentDefinition.BOM.BOMViews[1].ModelStateMemberName;

            var itemCounter = 1;
            await Task.Run(async () =>
            {
                foreach (var part in sheetMetalParts)
                {
                    if (_isCancelled) break;

                    var partData = await GetPartDataAsync(part.Key, part.Value, asmDoc.ComponentDefinition.BOM,
                        itemCounter++);
                    if (partData != null)
                    {
                        ((IProgress<PartData>)progress).Report(partData);
                        await Task.Delay(10);
                    }
                }
            });

            // Вызов проверки конфликтов после завершения сканирования
            CheckForPartNumberConflictsAsync(asmDoc.ComponentDefinition.BOM);
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                partCount = 1;
                partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                description = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description");

                var partData = await GetPartDataAsync(partNumber, 1, null, 1, partDoc);
                if (partData != null)
                {
                    ((IProgress<PartData>)progress).Report(partData);
                    await Task.Delay(10);
                }
            }
        }

        stopwatch.Stop();
        var elapsedTime = GetElapsedTime(stopwatch.Elapsed);

        if (int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
            UpdateQuantitiesWithMultiplier(multiplier);

        _isScanning = false;
        ProgressBar.IsIndeterminate = false;
        ProgressBar.Value = 0;
        ProgressLabel.Text = _isCancelled ? $"Статус: Прервано ({elapsedTime})" : $"Статус: Готово ({elapsedTime})";
        BottomPanel.Opacity = 1.0;
        UpdateDocumentInfo(documentType, partNumber, description, doc);
        UpdateFileCountLabel(_isCancelled ? 0 : partCount);

        ScanButton.IsEnabled = true;
        CancelButton.IsEnabled = false;
        ExportButton.IsEnabled = partCount > 0 && !_isCancelled;
        ClearButton.IsEnabled = _partsData.Count > 0;
        _lastScannedDocument = _isCancelled ? null : doc;

        if (isBackgroundMode) _thisApplication.UserInterfaceManager.UserInteractionDisabled = false;

        if (_isCancelled) MessageBoxHelper.ShowScanningCancelledInfo();
    }

    // Функция для проверки конфликта обозначений в BOM
    private async Task CheckForPartNumberConflictsAsync(BOM bom)
    {
        var partNumbers = new ConcurrentDictionary<string, List<(string PartNumber, string FileName)>>(); // Используем ConcurrentDictionary для потокобезопасности
        _conflictingParts.Clear(); // Очищаем список конфликтов перед началом проверки

        var tasks = new List<Task>(); // Список задач для параллельного выполнения

        // Проходим по каждому BOMView
        foreach (BOMView bomView in bom.BOMViews)
        {
            tasks.Add(Task.Run(() => ProcessBOMRowsForConflictsAsync(bomView, partNumbers)));
        }

        await Task.WhenAll(tasks); // Ожидаем завершения всех задач

        // Оставляем только те записи, у которых больше одного файла (т.е. есть конфликт)
        var conflictingPartNumbers = partNumbers.Where(p => p.Value.Count > 1).ToDictionary(p => p.Key, p => p.Value);

        if (conflictingPartNumbers.Any())
        {
            _conflictingParts.AddRange(conflictingPartNumbers.SelectMany(entry => entry.Value.Select(v => new PartData { PartNumber = v.PartNumber })));

            // Используем Dispatcher для обновления UI
            await Dispatcher.InvokeAsync(() =>
            {
                MessageBox.Show($"Обнаружены конфликты обозначений.\nОбщее количество конфликтов: {_conflictingParts.Count}\n\nОбнаружены различные модели с одинаковыми обозначениями, что может привести к ошибкам. Модели должны иметь уникальные обозначения.\n\nИспользуйте кнопку \"Анализ обозначений\" на панели инструментов для анализа данных.",
                    "Конфликт обозначений",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                ConflictFilesButton.IsEnabled = true; // Включаем кнопку для просмотра подробностей о конфликтах
                _conflictFileDetails = conflictingPartNumbers; // Сохраняем подробную информацию о конфликтах для вывода при нажатии кнопки
            });
        }
        else
        {
            await Dispatcher.InvokeAsync(() =>
            {
                ConflictFilesButton.IsEnabled = false; // Отключаем кнопку, если нет конфликтов
            });
        }
    }
    private async Task ProcessBOMRowsForConflictsAsync(BOMView bomView, ConcurrentDictionary<string, List<(string PartNumber, string FileName)>> partNumbers)
    {
        // Параллельно обрабатываем строки BOM
        var tasks = new List<Task>();

        foreach (BOMRow row in bomView.BOMRows)
        {
            // Исключаем библиотечные и приобретённые компоненты
            if (row.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure || row.BOMStructure == BOMStructureEnum.kPhantomBOMStructure)
            {
                continue; // Пропускаем библиотечные и приобретённые компоненты
            }

            foreach (ComponentDefinition componentDefinition in row.ComponentDefinitions)
            {
                tasks.Add(Task.Run(() =>
                {
                    if (_isCancelled) return; // Прерываем выполнение, если операция была отменена

                    var document = componentDefinition.Document as Document;

                    if (document != null)
                    {
                        if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                        {
                            var partDocument = document as PartDocument;

                            // Проверяем, что это листовая деталь
                            if (partDocument != null && partDocument.ComponentDefinition is SheetMetalComponentDefinition)
                            {
                                var partNumber = GetProperty(partDocument.PropertySets["Design Tracking Properties"], "Part Number");
                                var fileName = partDocument.FullFileName;

                                // Если обозначение уже существует, проверяем на конфликт
                                partNumbers.AddOrUpdate(partNumber,
                                    new List<(string PartNumber, string FileName)> { (partNumber, fileName) },
                                    (key, existingList) =>
                                    {
                                        if (!existingList.Any(p => p.FileName == fileName))
                                        {
                                            existingList.Add((partNumber, fileName)); // Добавляем файл к существующему конфликту
                                        }
                                        return existingList;
                                    });
                            }
                        }
                        else if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                        {
                            var asmDoc = document as AssemblyDocument;
                            if (asmDoc != null)
                            {
                                foreach (BOMView subBomView in asmDoc.ComponentDefinition.BOM.BOMViews)
                                {
                                    ProcessBOMRowsForConflictsAsync(subBomView, partNumbers).Wait();
                                }
                            }
                        }
                    }
                }));
            }
        }

        await Task.WhenAll(tasks); // Ожидаем завершения всех задач
    }
    private void ConflictFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_conflictFileDetails == null || _conflictFileDetails.Count == 0)
        {
            MessageBox.Show("Нет конфликтующих обозначений для отображения.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Создаем и показываем окно с деталями конфликтов
        var conflictWindow = new ConflictDetailsWindow(_conflictFileDetails);
        conflictWindow.Owner = this; // Устанавливаем основное окно как родительское
        conflictWindow.ShowDialog(); // Ожидаем закрытия окна
    }

    private void UpdateSubfolderOptions()
    {
        // Опция "В папку с деталью" - чекбокс и текстовое поле должны быть отключены
        if (PartFolderRadioButton.IsChecked == true)
        {
            EnableSubfolderCheckBox.IsEnabled = false;
            EnableSubfolderCheckBox.IsChecked = false;
            SubfolderNameTextBox.IsEnabled = false;
        }
        else
        {
            // Для всех остальных опций чекбокс активен
            EnableSubfolderCheckBox.IsEnabled = true;
            SubfolderNameTextBox.IsEnabled = EnableSubfolderCheckBox.IsChecked == true;
        }
    }

    private void SubfolderNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var forbiddenChars = "\\/:*?\"<>|";
        var textBox = sender as TextBox;
        var caretIndex = textBox.CaretIndex;

        foreach (var ch in forbiddenChars)
            if (textBox.Text.Contains(ch))
            {
                textBox.Text = textBox.Text.Replace(ch.ToString(), string.Empty);
                caretIndex--; // Уменьшаем позицию курсора на 1
            }

        // Возвращаем курсор на правильное место после удаления символов
        textBox.CaretIndex = Math.Max(caretIndex, 0);
    }
    // Метод для получения и установки информации о проекте
    private void SetProjectFolderInfo()
    {
        try
        {
            // Получаем текущий активный проект
            var activeProject = _thisApplication.DesignProjectManager.ActiveDesignProject;

            // Извлекаем название и путь проекта
            var projectName = activeProject.Name;
            var projectWorkspacePath = activeProject.WorkspacePath;

            // Обновляем текст радиокнопки
            //projectFolderRadioButton.Content = $"Папка проекта [{projectName}] {projectWorkspacePath}";

            // Обновляем TextBlock с названием проекта
            this.ProjectName.Text = $"Проект: {projectName}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Не удалось получить информацию о проекте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // Вызываем при инициализации приложения
    private void InitializeProjectFolder()
    {
        SetProjectFolderInfo(); // Устанавливаем информацию о проекте
    }

    // Обновленный метод для обработки нажатия радиокнопки
    private void ProjectFolderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        SetProjectFolderInfo(); // Устанавливаем информацию о проекте при выборе радиокнопки
        UpdateSubfolderOptions();
    }
    private void PartFolderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateSubfolderOptions();
    }

    private void ComponentFolderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateSubfolderOptions();
    }

    private void ChooseFolderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateSubfolderOptions();
    }

    private void FixedFolderRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateSubfolderOptions();
    }

    private void EnableSubfolderCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        SubfolderNameTextBox.IsEnabled = true;
    }

    private void EnableSubfolderCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        SubfolderNameTextBox.IsEnabled = false;
        SubfolderNameTextBox.Text = string.Empty;
    }

    private string GetProperty(PropertySet propertySet, string propertyName)
    {
        try
        {
            var property = propertySet[propertyName];
            return property.Value.ToString();
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    private void ReadIPropertyValuesFromPart(PartDocument partDoc, PartData partData)
    {
        var propertySets = partDoc.PropertySets;

        // Проверка и чтение свойства "Author"
        if (IsColumnPresent("Автор")) partData.Author = GetProperty(propertySets["Summary Information"], "Author");

        // Проверка и чтение свойства "Revision"
        if (IsColumnPresent("Ревизия"))
            partData.Revision = GetProperty(propertySets["Design Tracking Properties"], "Revision Number");

        // Проверка и чтение свойства "Project"
        if (IsColumnPresent("Проект")) partData.Project = GetProperty(propertySets["Summary Information"], "Project");

        // Проверка и чтение свойства "Stock number"
        if (IsColumnPresent("Инвентарный номер"))
            partData.StockNumber = GetProperty(propertySets["Design Tracking Properties"], "Stock Number");
    }

    private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, BOM bom, int itemNumber,
        PartDocument? partDoc = null)
    {
        // Открываем документ, если он не передан
        if (partDoc == null)
        {
            partDoc = OpenPartDocument(partNumber);
            if (partDoc == null) return null;
        }

        // Создаем новый объект PartData и заполняем его основными свойствами
        var partData = new PartData
        {
            PartNumber = await GetPropertyExpressionOrValueAsync(partDoc, "Part Number"),
            PartNumberExpression = await GetPropertyExpressionOrValueAsync(partDoc, "Part Number", true),
            Description = await GetPropertyExpressionOrValueAsync(partDoc, "Description"),
            DescriptionExpression = await GetPropertyExpressionOrValueAsync(partDoc, "Description", true),
            Material = GetMaterialForPart(partDoc),
            Thickness = GetThicknessForPart(partDoc).ToString("F1") + " мм", // Получаем толщину
            ModelState = partDoc.ModelStateName,
            OriginalQuantity = quantity,
            Quantity = quantity,
            QuantityColor = Brushes.Black,
            Item = itemNumber
        };

        // Получаем выражения и значения для пользовательских свойств
        foreach (var customProperty in _customPropertiesList)
        {
            // Асинхронно получаем значения пользовательских свойств
            partData.CustomProperties[customProperty] =
                await GetPropertyExpressionOrValueAsync(partDoc, customProperty);
            partData.CustomPropertyExpressions[customProperty] =
                await GetPropertyExpressionOrValueAsync(partDoc, customProperty, true);
        }

        // Асинхронно получаем превью-изображение (если доступно)
        partData.Preview = await GetThumbnailAsync(partDoc);

        // Проверяем наличие развертки для листовых деталей и задаем цвет
        var smCompDef = partDoc.ComponentDefinition as SheetMetalComponentDefinition;
        var hasFlatPattern = smCompDef != null && smCompDef.HasFlatPattern;
        partData.FlatPatternColor = hasFlatPattern ? Brushes.Green : Brushes.Red;

        // Чтение значений iProperty для Автор, Ревизия, Проект, Инвентарный номер
        ReadIPropertyValuesFromPart(partDoc, partData);

        return partData;
    }

    private BitmapImage GenerateDxfThumbnail(string dxfDirectory, string partNumber)
    {
        var searchPattern = partNumber + "*.dxf"; // Шаблон поиска
        var dxfFiles = Directory.GetFiles(dxfDirectory, searchPattern);

        if (dxfFiles.Length == 0) return null;

        try
        {
            var dxfFilePath = dxfFiles[0]; // Берем первый найденный файл, соответствующий шаблону
            var generator = new DxfThumbnailGenerator.DxfThumbnailGenerator();
            var bitmap = generator.GenerateThumbnail(dxfFilePath);

            BitmapImage bitmapImage = null;

            // Инициализация изображения должна выполняться в UI потоке
            Dispatcher.Invoke(() =>
            {
                using (var memoryStream = new MemoryStream())
                {
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Position = 0;

                    bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memoryStream;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                }
            });

            return bitmapImage;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Ошибка при генерации миниатюры DXF: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
        }
    }

    private string GetModelStateName(PartDocument partDoc)
    {
        try
        {
            var modelStates = partDoc.ComponentDefinition.ModelStates;
            if (modelStates.Count > 0)
            {
                var activeStateName = modelStates.ActiveModelState.Name;
                return activeStateName;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при получении Model State: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        return "Ошибка получения имени";
    }

    private string GetModelStateName(AssemblyDocument asmDoc)
    {
        try
        {
            var modelStates = asmDoc.ComponentDefinition.ModelStates;
            if (modelStates.Count > 0)
            {
                var activeStateName = modelStates.ActiveModelState.Name;
                return activeStateName;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при получении Model State: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        return "Ошибка получения имени";
    }

    private async Task<BitmapImage> GetThumbnailAsync(PartDocument document)
    {
        try
        {
            BitmapImage bitmap = null;
            await Dispatcher.InvokeAsync(() =>
            {
                var apprentice = new ApprenticeServerComponent();
                var apprenticeDoc = apprentice.Open(document.FullDocumentName);

                var thumbnail = apprenticeDoc.Thumbnail;
                var img = IPictureDispConverter.PictureDispToImage(thumbnail);

                if (img != null)
                    using (var memoryStream = new MemoryStream())
                    {
                        img.Save(memoryStream, ImageFormat.Png);
                        memoryStream.Position = 0;

                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = memoryStream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                    }
            });

            return bitmap;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show("Ошибка при получении миниатюры: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
        }
    }

    private void ProcessComponentOccurrences(ComponentOccurrences occurrences, Dictionary<string, int> sheetMetalParts)
    {
        foreach (ComponentOccurrence occ in occurrences)
        {
            if (_isCancelled) break;

            if (occ.Suppressed) continue;

            try
            {
                if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var partDoc = occ.Definition.Document as PartDocument;
                    if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" &&
                        occ.BOMStructure != BOMStructureEnum.kReferenceBOMStructure)
                    {
                        var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                sheetMetalParts[partNumber]++;
                            else
                                sheetMetalParts.Add(partNumber, 1);
                        }
                    }
                }
                else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ProcessComponentOccurrences((ComponentOccurrences)occ.SubOccurrences, sheetMetalParts);
                }
            }
            catch (COMException ex)
            {
                // Обработка ошибок COM
            }
            catch (Exception ex)
            {
                // Обработка общих ошибок
            }
        }
    }

    private void ProcessBOM(BOM bom, Dictionary<string, int> sheetMetalParts)
    {
        BOMView? bomView = null;
        foreach (BOMView view in bom.BOMViews)
        {
            bomView = view;
            break;
        }

        if (bomView == null)
        {
            MessageBox.Show("Не найдено ни одного представления спецификации.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        foreach (BOMRow row in bomView.BOMRows)
        {
            if (_isCancelled) break;

            try
            {
                var componentDefinition = row.ComponentDefinitions[1];
                if (componentDefinition == null) continue;

                var document = componentDefinition.Document as Document;
                if (document == null) continue;

                if (row.BOMStructure == BOMStructureEnum.kReferenceBOMStructure ||
                    row.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure) continue;

                if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var partDoc = document as PartDocument;
                    if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" &&
                        row.BOMStructure != BOMStructureEnum.kReferenceBOMStructure)
                    {
                        var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                sheetMetalParts[partNumber] += row.ItemQuantity;
                            else
                                sheetMetalParts.Add(partNumber, row.ItemQuantity);
                        }
                    }
                }
                else if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    var asmDoc = document as AssemblyDocument;
                    if (asmDoc != null) ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts);
                }
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80070057))
            {
                // Обработка ошибок COM
            }
            catch (Exception ex)
            {
                // Обработка общих ошибок
            }
        }
    }

    private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _fixedFolderPath = dialog.SelectedPath;

            if (_fixedFolderPath.Length > 40)
                // Гарантируем, что показываем именно последние 40 символов
                FixedFolderPathTextBlock.Text = "..." + _fixedFolderPath.Substring(_fixedFolderPath.Length - 40);
            else
                FixedFolderPathTextBlock.Text = _fixedFolderPath;
        }
    }

private bool PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch)
{
    stopwatch = new Stopwatch();
    targetDir = string.Empty;
    multiplier = 1; // Присваиваем значение по умолчанию

    // Обрабатываем выбор радио-кнопки
    if (ChooseFolderRadioButton.IsChecked == true)
    {
        // Открываем диалог выбора папки
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            targetDir = dialog.SelectedPath;
        else
            return false;
    }
    else if (ComponentFolderRadioButton.IsChecked == true)
    {
        // Папка компонента
        targetDir = Path.GetDirectoryName(_thisApplication.ActiveDocument.FullFileName);
    }
    else if (FixedFolderRadioButton.IsChecked == true)
    {
        if (string.IsNullOrEmpty(_fixedFolderPath))
        {
            MessageBox.Show("Выберите фиксированную папку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        targetDir = _fixedFolderPath;
    }
    else if (ProjectFolderRadioButton.IsChecked == true) // Новый код для выбора папки проекта
    {
        try
        {
            targetDir = _thisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            if (string.IsNullOrEmpty(targetDir))
            {
                MessageBox.Show("Не удалось получить путь проекта.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при получении папки проекта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    if (EnableSubfolderCheckBox.IsChecked == true && !string.IsNullOrEmpty(SubfolderNameTextBox.Text))
    {
        targetDir = Path.Combine(targetDir, SubfolderNameTextBox.Text);
        if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
    }

    if (IncludeQuantityInFileNameCheckBox.IsChecked == true && !int.TryParse(MultiplierTextBox.Text, out multiplier))
    {
        MessageBox.Show("Введите допустимое целое число для множителя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return false;
    }

    CancelButton.IsEnabled = true;
    _isCancelled = false;

    ProgressBar.IsIndeterminate = false;
    ProgressBar.Value = 0;
    ProgressLabel.Text = "Статус: Экспорт данных...";

    stopwatch = Stopwatch.StartNew();

    return true;
}

    private void FinalizeExport(bool isCancelled, Stopwatch stopwatch, int processedCount, int skippedCount)
    {
        stopwatch.Stop();
        var elapsedTime = GetElapsedTime(stopwatch.Elapsed);

        CancelButton.IsEnabled = false;
        ScanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
        ClearButton.IsEnabled = _partsData.Count > 0; // Активируем кнопку "Очистить" после завершения экспорта

        if (isCancelled)
        {
            ProgressLabel.Text = $"Статус: Прервано ({elapsedTime})";
            MessageBoxHelper.ShowExportCancelledInfo();
        }
        else
        {
            ProgressLabel.Text = $"Статус: Завершено ({elapsedTime})";
            MessageBoxHelper.ShowExportCompletedInfo(processedCount, skippedCount, elapsedTime);
        }
    }

    private async Task ExportWithoutScan()
    {
        // Очистка таблицы перед началом скрытого экспорта
        ClearList_Click(this, null);

        if (_thisApplication == null)
        {
            MessageBoxHelper.ShowInventorNotRunningError();
            return;
        }

        Document doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBoxHelper.ShowNoDocumentOpenWarning();
            return;
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch)) return;

        ClearButton.IsEnabled = false;
        ScanButton.IsEnabled = false;

        var processedCount = 0;
        var skippedCount = 0;

        // Скрытое сканирование
        var sheetMetalParts = new Dictionary<string, int>();

        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)doc;
            if (TraverseRadioButton.IsChecked == true)
                await Task.Run(() =>
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));

            // Создаем временный список PartData для экспорта
            var tempPartsDataList = new List<PartData>();
            foreach (var part in sheetMetalParts)
                tempPartsDataList.Add(new PartData
                {
                    PartNumber = part.Key,
                    OriginalQuantity = part.Value,
                    Quantity = part.Value * multiplier
                });

            await Task.Run(() =>
                ExportDXF(tempPartsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                    false)); // false - без генерации миниатюр
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                var tempPartData = new List<PartData>
                {
                    new()
                    {
                        PartNumber = partNumber,
                        OriginalQuantity = 1,
                        Quantity = multiplier
                    }
                };
                await Task.Run(() =>
                    ExportDXF(tempPartData, targetDir, multiplier, ref processedCount, ref skippedCount,
                        false)); // false - без генерации миниатюр
            }
            else
            {
                MessageBoxHelper.ShowNotSheetMetalWarning();
            }
        }

        ClearButton.IsEnabled = true;
        ScanButton.IsEnabled = true;

        // Здесь мы НЕ обновляем fileCountLabelBottom
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);

        // Деактивируем кнопку "Экспорт" после завершения скрытого экспорта
        ExportButton.IsEnabled = false;
    }

    private void ResetEditMode()
    {
        // Сбрасываем режим редактирования
        if (_isEditing)
        {
            _isEditing = false; // Отключаем режим редактирования
            PartsDataGrid.IsReadOnly = true; // Делаем таблицу нередактируемой
            EditButton.Content = "Редактировать"; // Меняем текст кнопки на "Редактировать"
            SaveButton.IsEnabled = false; // Деактивируем кнопку сохранения

            // Обновляем IsReadOnly для каждого элемента данных
            foreach (var item in
                     _partsData) item.IsReadOnly = true; // Устанавливаем флаг, что элемент теперь нередактируемый

            // Обновляем отображение таблицы
            PartsDataGrid.Items.Refresh(); // Обновляем интерфейс, чтобы стиль ячеек обновился
        }
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        ResetEditMode(); // Сброс режима редактирования перед экспортом

        if (_isCtrlPressed)
        {
            await ExportWithoutScan();
            return;
        }

        if (_thisApplication == null)
        {
            MessageBoxHelper.ShowInventorNotRunningError();
            return;
        }

        var isBackgroundMode = BackgroundModeCheckBox.IsChecked == true;
        if (isBackgroundMode) _thisApplication.UserInterfaceManager.UserInteractionDisabled = true;

        Document? doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBoxHelper.ShowNoDocumentOpenWarning();
            return;
        }

        if (doc != _lastScannedDocument)
        {
            MessageBox.Show("Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.",
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            MultiplierTextBox.Text = "1";
            return;
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch))
        {
            if (isBackgroundMode) _thisApplication.UserInterfaceManager.UserInteractionDisabled = false;
            return;
        }

        ClearButton.IsEnabled = false;
        ScanButton.IsEnabled = false;

        var processedCount = 0;
        var skippedCount = 0;

        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)doc;
            var sheetMetalParts = new Dictionary<string, int>();
            if (TraverseRadioButton.IsChecked == true)
                await Task.Run(() =>
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));

            await ExportDXFFiles(sheetMetalParts, targetDir, multiplier, stopwatch);
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                await ExportSinglePartDXF(partDoc, targetDir, multiplier, stopwatch);
            else
                MessageBoxHelper.ShowNotSheetMetalWarning();
        }

        ClearButton.IsEnabled = true;
        ScanButton.IsEnabled = true;

        if (isBackgroundMode) _thisApplication.UserInterfaceManager.UserInteractionDisabled = false;

        ExportButton.IsEnabled = true;
    }

    private string GetElapsedTime(TimeSpan timeSpan)
    {
        if (timeSpan.TotalMinutes >= 1)
            return $"{(int)timeSpan.TotalMinutes} мин. {timeSpan.Seconds} сек.";
        return $"{timeSpan.Seconds} сек.";
    }

    private async void ExportSelectedDXF_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для экспорта.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var itemsWithoutFlatPattern = selectedItems.Where(p => p.FlatPatternColor == Brushes.Red).ToList();
        if (itemsWithoutFlatPattern.Count == selectedItems.Count)
        {
            MessageBoxHelper.ShowNoFlatPatternWarning();
            return;
        }

        if (itemsWithoutFlatPattern.Count > 0)
        {
            var result =
                MessageBox.Show("Некоторые выбранные файлы не содержат разверток и будут пропущены. Продолжить?",
                    "Информация", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.No) return;

            selectedItems = selectedItems.Except(itemsWithoutFlatPattern).ToList();
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch)) return;

        ClearButton.IsEnabled = false; // Деактивируем кнопку "Очистить" в процессе экспорта
        ScanButton.IsEnabled = false; // Деактивируем кнопку "Сканировать" в процессе экспорта

        var processedCount = 0;
        var skippedCount = itemsWithoutFlatPattern.Count;

        await Task.Run(() =>
        {
            ExportDXF(selectedItems, targetDir, multiplier, ref processedCount, ref skippedCount,
                true); // true - с генерацией миниатюр
        });

        UpdateFileCountLabel(selectedItems.Count);

        ClearButton.IsEnabled = true; // Активируем кнопку "Очистить" после завершения экспорта
        ScanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private void OverrideQuantity_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для переопределения количества.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new OverrideQuantityDialog();
        dialog.Owner = this;
        dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

        if (dialog.ShowDialog() == true && dialog.NewQuantity.HasValue)
        {
            foreach (var item in selectedItems)
            {
                item.Quantity = dialog.NewQuantity.Value;
                item.IsOverridden = true;
                item.QuantityColor = Brushes.Red; // Переопределенное количество окрашивается в красный цвет
            }

            PartsDataGrid.Items.Refresh();
        }
    }

    private void ExportDXF(IEnumerable<PartData> partsDataList, string targetDir, int multiplier,
        ref int processedCount, ref int skippedCount, bool generateThumbnails)
    {
        var organizeByMaterial = false;
        var organizeByThickness = false;
        var includeQuantityInFileName = false;

        // Выполняем доступ к элементам управления через Dispatcher
        Dispatcher.Invoke(() =>
        {
            organizeByMaterial = OrganizeByMaterialCheckBox.IsChecked == true;
            organizeByThickness = OrganizeByThicknessCheckBox.IsChecked == true;
            includeQuantityInFileName = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        });

        var totalParts = partsDataList.Count();

        Dispatcher.Invoke(() =>
        {
            ProgressBar.Maximum = totalParts;
            ProgressBar.Value = 0;
        });

        var localProcessedCount = processedCount;
        var localSkippedCount = skippedCount;

        foreach (var partData in partsDataList)
        {
            if (_isCancelled) break;

            var partNumber = partData.PartNumber;
            var qty = partData.IsOverridden ? partData.Quantity : partData.OriginalQuantity * multiplier;

            // Доступ к RadioButton через Dispatcher для потока безопасности
            Dispatcher.Invoke(() =>
            {
                if (PartFolderRadioButton.IsChecked == true)
                {
                    targetDir = GetPartDocumentFullPath(partNumber);
                    targetDir = Path.GetDirectoryName(targetDir);
                }
            });

            PartDocument partDoc = null;
            var isPartDocOpenedHere = false;
            try
            {
                partDoc = OpenPartDocument(partNumber);
                if (partDoc == null) throw new Exception("Файл детали не найден или не может быть открыт");

                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var material = GetMaterialForPart(partDoc);
                var thickness = GetThicknessForPart(partDoc);

                var materialDir = organizeByMaterial ? Path.Combine(targetDir, material) : targetDir;
                if (!Directory.Exists(materialDir)) Directory.CreateDirectory(materialDir);

                var thicknessDir = organizeByThickness
                    ? Path.Combine(materialDir, thickness.ToString("F1"))
                    : materialDir;
                if (!Directory.Exists(thicknessDir)) Directory.CreateDirectory(thicknessDir);

                var filePath = Path.Combine(thicknessDir,
                    partNumber + (includeQuantityInFileName ? " - " + qty : "") + ".dxf");

                if (!IsValidPath(filePath)) continue;

                var exportSuccess = false;

                if (smCompDef.HasFlatPattern)
                {
                    var flatPattern = smCompDef.FlatPattern;
                    var oDataIO = flatPattern.DataIO;

                    try
                    {
                        string options = null;
                        Dispatcher.Invoke(() => PrepareExportOptions(out options));

                        // Интеграция настроек слоев
                        var layerOptionsBuilder = new StringBuilder();
                        var invisibleLayersBuilder = new StringBuilder();

                        foreach (var layer in LayerSettings) // Используем настройки слоев
                        {
                            if (layer.HasVisibilityOption && !layer.IsVisible)
                            {
                                invisibleLayersBuilder.Append($"{layer.LayerName};");
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(layer.CustomName))
                                layerOptionsBuilder.Append($"&{layer.DisplayName}={layer.CustomName}");

                            if (layer.SelectedLineType != "Default")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}LineType={LayerSettingsApp.MainWindow.GetLineTypeValue(layer.SelectedLineType)}");

                            if (layer.SelectedColor != "White")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}Color={LayerSettingsApp.MainWindow.GetColorValue(layer.SelectedColor)}");
                        }

                        if (invisibleLayersBuilder.Length > 0)
                        {
                            invisibleLayersBuilder.Length -= 1; // Убираем последний символ ";"
                            layerOptionsBuilder.Append($"&InvisibleLayers={invisibleLayersBuilder}");
                        }

                        // Объединяем общие параметры и параметры слоев
                        options += layerOptionsBuilder.ToString();

                        var dxfOptions = $"FLAT PATTERN DXF?{options}";

                        // Debugging output
                        Debug.WriteLine($"DXF Options: {dxfOptions}");

                        // Создаем файл, если он не существует
                        if (!File.Exists(filePath)) File.Create(filePath).Close();

                        // Проверка на занятость файла
                        if (IsFileLocked(filePath))
                        {
                            var result = MessageBox.Show(
                                $"Файл {filePath} занят другим приложением. Прервать операцию?",
                                "Предупреждение",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning,
                                MessageBoxResult.No // Устанавливаем "No" как значение по умолчанию
                            );

                            if (result == MessageBoxResult.Yes)
                            {
                                _isCancelled = true; // Прерываем операцию
                                break;
                            }

                            while (IsFileLocked(filePath))
                            {
                                var waitResult = MessageBox.Show(
                                    "Ожидание разблокировки файла. Возможно файл открыт в программе просмотра. Закройте файл и нажмите OK.",
                                    "Информация",
                                    MessageBoxButton.OKCancel,
                                    MessageBoxImage.Information,
                                    MessageBoxResult.Cancel // Устанавливаем "Cancel" как значение по умолчанию
                                );

                                if (waitResult == MessageBoxResult.Cancel)
                                {
                                    _isCancelled = true; // Прерываем операцию
                                    break;
                                }

                                Thread.Sleep(1000); // Ожидание 1 секунды перед повторной проверкой
                            }

                            if (_isCancelled) break;
                        }

                        oDataIO.WriteDataToFile(dxfOptions, filePath);
                        exportSuccess = true;
                    }
                    catch (Exception ex)
                    {
                        // Обработка ошибок экспорта DXF
                        Debug.WriteLine($"Ошибка экспорта DXF: {ex.Message}");
                    }
                }

                BitmapImage dxfPreview = null;
                if (generateThumbnails) dxfPreview = GenerateDxfThumbnail(thicknessDir, partNumber);

                Dispatcher.Invoke(() =>
                {
                    partData.ProcessingColor = exportSuccess ? Brushes.Green : Brushes.Red;
                    partData.DxfPreview = dxfPreview;
                    PartsDataGrid.Items.Refresh();
                });

                if (exportSuccess)
                    localProcessedCount++;
                else
                    localSkippedCount++;

                Dispatcher.Invoke(() => { ProgressBar.Value = Math.Min(localProcessedCount, ProgressBar.Maximum); });
            }
            catch (Exception ex)
            {
                localSkippedCount++;
                Debug.WriteLine($"Ошибка обработки детали: {ex.Message}");
            }
        }

        processedCount = localProcessedCount;
        skippedCount = localSkippedCount;

        Dispatcher.Invoke(() =>
        {
            ProgressBar.Value = ProgressBar.Maximum; // Установка значения 100% по завершению
        });
    }

    private bool IsFileLocked(string filePath)
    {
        try
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                stream.Close();
            }
        }
        catch (IOException)
        {
            return true;
        }

        return false;
    }

    private async Task ExportDXFFiles(Dictionary<string, int> sheetMetalParts, string targetDir, int multiplier,
        Stopwatch stopwatch)
    {
        var processedCount = 0;
        var skippedCount = 0;

        var partsDataList = _partsData.Where(p => sheetMetalParts.ContainsKey(p.PartNumber)).ToList();
        await Task.Run(
            () => ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                true)); // true - с генерацией миниатюр

        UpdateFileCountLabel(partsDataList.Count);
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private async Task ExportSinglePartDXF(PartDocument partDoc, string targetDir, int multiplier, Stopwatch stopwatch)
    {
        var processedCount = 0;
        var skippedCount = 0;

        var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
        var partData = _partsData.FirstOrDefault(p => p.PartNumber == partNumber);

        if (partData != null)
        {
            var partsDataList = new List<PartData> { partData };
            await Task.Run(() =>
                ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                    true)); // true - с генерацией миниатюр
            UpdateFileCountLabel(1);
        }

        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private PartDocument OpenPartDocument(string partNumber)
    {
        var docs = _thisApplication.Documents;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd &&
                GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                return pd; // Возвращаем найденный открытый документ

        // Если документ не найден
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private string GetMaterialForPart(PartDocument partDoc)
    {
        try
        {
            return partDoc.ComponentDefinition.Material.Name;
        }
        catch (Exception)
        {
            // Обработка ошибок получения материала
        }

        return "Ошибка получения имени";
    }

    private double GetThicknessForPart(PartDocument partDoc)
    {
        try
        {
            var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
            var thicknessParam = smCompDef.Thickness;
            return
                Math.Round((double)thicknessParam.Value * 10,
                    1); // Извлекаем значение толщины и переводим в мм, округляем до одной десятой
        }
        catch (Exception)
        {
            // Обработка ошибок получения толщины
        }

        return 0;
    }

    private bool IsValidPath(string path)
    {
        try
        {
            new FileInfo(path);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void IncludeQuantityInFileNameCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        MultiplierTextBox.IsEnabled = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        if (!MultiplierTextBox.IsEnabled)
        {
            MultiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);
        }
    }

    private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
        {
            UpdateQuantitiesWithMultiplier(multiplier);

            // Проверка на null перед изменением видимости кнопки
            if (ClearMultiplierButton != null)
                ClearMultiplierButton.Visibility = multiplier > 1 ? Visibility.Visible : Visibility.Collapsed;
        }
        else
        {
            // Если введенное значение некорректное, сбрасываем текст на "1" и скрываем кнопку сброса
            MultiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);

            if (ClearMultiplierButton != null) ClearMultiplierButton.Visibility = Visibility.Collapsed;
        }
    }

    private void ClearMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        // Сбрасываем множитель на 1
        MultiplierTextBox.Text = "1";
        UpdateQuantitiesWithMultiplier(1);

        // Скрываем кнопку сброса
        ClearMultiplierButton.Visibility = Visibility.Collapsed;
    }

    private void UpdateFileCountLabel(int count)
    {
        FileCountLabelBottom.Text = $"Найдено листовых деталей: {count}";
    }


    private void ResetProgressBar()
    {
        ProgressBar.IsIndeterminate = false;
        ProgressBar.Minimum = 0;
        ProgressBar.Maximum = 100;
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Статус: ";
        UpdateFileCountLabel(0); // Использование метода для сброса счетчика
        BottomPanel.Opacity = 0.75;
        DocumentInfoLabel.Text = string.Empty;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _isCancelled = true;
        CancelButton.IsEnabled = false;

        if (!_isScanning)
        {
            // If not scanning, reset the UI
            ResetProgressBar();
            BottomPanel.Opacity = 0.75;
            DocumentInfoLabel.Text = "";
            ExportButton.IsEnabled = false;
            ClearButton.IsEnabled = _partsData.Count > 0; // Обновляем состояние кнопки "Очистить"
            _lastScannedDocument = null;
            UpdateFileCountLabel(0);
        }
    }

    private void UpdateDocumentInfo(string documentType, string partNumber, string description, Document doc)
    {
        if (string.IsNullOrEmpty(documentType) && string.IsNullOrEmpty(partNumber) && string.IsNullOrEmpty(description))
        {
            DocumentInfoLabel.Text = string.Empty;
            ModelStateInfoRunBottom.Text = string.Empty; // Обновленное имя элемента
            BottomPanel.Opacity = 0.75;
            UpdateFileCountLabel(0); // Использование метода для сброса счетчика
            return;
        }

        var modelStateInfo = doc is PartDocument partDoc ? GetModelStateName(partDoc) :
            doc is AssemblyDocument asmDoc ? GetModelStateName(asmDoc) :
            "Ошибка получения имени";

        DocumentInfoLabel.Text = $"Тип документа: {documentType}\nОбозначение: {partNumber}\nОписание: {description}";

        if (modelStateInfo == "[Primary]" || modelStateInfo == "[Основной]")
        {
            ModelStateInfoRunBottom.Text = "[Основной]";
            ModelStateInfoRunBottom.Foreground =
                new SolidColorBrush(Colors.Black); // Цвет текста - черный (по умолчанию)
        }
        else
        {
            ModelStateInfoRunBottom.Text = modelStateInfo;
            ModelStateInfoRunBottom.Foreground = new SolidColorBrush(Colors.Red); // Цвет текста - красный
        }

        UpdateFileCountLabel(_partsData.Count); // Использование метода для обновления счетчика
        BottomPanel.Opacity = 1.0;
    }

    private void ClearList_Click(object sender, RoutedEventArgs e)
    {
        // Очищаем данные в таблице или любом другом источнике данных
        _partsData.Clear();
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Статус: ";
        DocumentInfoLabel.Text = string.Empty;
        ModelStateInfoRunBottom.Text = string.Empty;
        BottomPanel.Opacity = 0.5;
        ExportButton.IsEnabled = false;
        ClearButton.IsEnabled = false; // Делаем кнопку "Очистить" неактивной после очистки

        // Обнуляем информацию о документе
        UpdateDocumentInfo(string.Empty, string.Empty, string.Empty, null);
        UpdateFileCountLabel(0); // Сброс счетчика файлов

        // Отключаем кнопку "Анализ обозначений" и очищаем список конфликтов
        ConflictFilesButton.IsEnabled = false;
        _conflictFileDetails.Clear(); // Очищаем список с конфликтами (или используем нужную коллекцию)

        // Обновляем состояние кнопки после очистки данных
        UpdateEditButtonState();
    }

    private void partsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_partsData.Count == 0) e.Handled = true; // Предотвращаем открытие контекстного меню, если таблица пуста
    }

    private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для открытия расположения файла.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = GetPartDocumentFullPath(partNumber);

            // Проверка на null перед использованием fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Проверка существования файла
            if (File.Exists(fullPath))
                try
                {
                    var argument = "/select, \"" + fullPath + "\"";
                    Process.Start("explorer.exe", argument);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии проводника для файла по пути {fullPath}: {ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            else
                MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSelectedModels_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для открытия моделей.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = GetPartDocumentFullPath(partNumber);

            // Проверка на null перед использованием fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Проверка существования файла
            if (File.Exists(fullPath))
                try
                {
                    // Открытие документа в Inventor
                    _thisApplication.Documents.Open(fullPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии файла по пути {fullPath}: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            else
                MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string GetPartDocumentFullPath(string partNumber)
    {
        var docs = _thisApplication.Documents;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd &&
                GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                return pd.FullFileName; // Возвращаем полный путь найденного документа

        // Если документ не найден среди открытых
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private void partsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Активация пункта "Редактировать iProperty" только при выборе одной строки
        EditIPropertyMenuItem.IsEnabled = PartsDataGrid.SelectedItems.Count == 1;
    }

    private async void AddCustomIPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        var inputDialog = new CustomIPropertyDialog();
        if (inputDialog.ShowDialog() == true)
        {
            var customPropertyName = inputDialog.CustomPropertyName;

            // Проверка на существование в списке пользовательских свойств
            if (_customPropertiesList.Contains(customPropertyName))
            {
                MessageBox.Show($"Свойство '{customPropertyName}' уже добавлено.", "Предупреждение",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Проверка на существование колонки
            if (PartsDataGrid.Columns.Any(c => c.Header as string == customPropertyName))
            {
                MessageBox.Show($"Столбец с именем '{customPropertyName}' уже существует.", "Предупреждение",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _customPropertiesList.Add(customPropertyName);

            // Добавляем новую колонку в DataGrid
            AddCustomIPropertyColumn(customPropertyName);

            // Если включен фоновый режим, отключаем взаимодействие с UI
            if (BackgroundModeCheckBox.IsChecked == true)
                _thisApplication.UserInterfaceManager.UserInteractionDisabled = true;

            try
            {
                // Заполняем значения Custom iProperty для уже существующих деталей асинхронно
                await FillCustomPropertyAsync(customPropertyName);
            }
            finally
            {
                // Восстанавливаем взаимодействие с UI после завершения асинхронной операции
                if (BackgroundModeCheckBox.IsChecked == true)
                    _thisApplication.UserInterfaceManager.UserInteractionDisabled = false;
            }
        }
    }

    private async Task FillCustomPropertyAsync(string customPropertyName)
    {
        foreach (var partData in _partsData)
        {
            // Получаем значение custom iProperty асинхронно
            var propertyValue = await Task.Run(() => GetCustomIPropertyValue(partData.PartNumber, customPropertyName));

            // Добавляем значение в коллекцию свойств для конкретного элемента
            partData.AddCustomProperty(customPropertyName, propertyValue);

            // Обновляем UI
            PartsDataGrid.Items.Refresh();

            // Проверяем, включен ли фоновый режим
            if (BackgroundModeCheckBox.IsChecked == true)
                // Дополнительная задержка для плавности и снижения нагрузки на UI
                await Task.Delay(50);
        }
    }

    private void RemoveCustomIPropertyColumn(string propertyName)
    {
        // Удаление столбца из DataGrid
        var columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.Header as string == propertyName);
        if (columnToRemove != null) PartsDataGrid.Columns.Remove(columnToRemove);

        // Удаление всех данных, связанных с этим Custom IProperty
        foreach (var partData in _partsData)
            if (partData.CustomProperties.ContainsKey(propertyName))
                partData.RemoveCustomProperty(propertyName);

        // Обновляем список доступных свойств
        _customPropertiesList.Remove(propertyName); // Не забудьте удалить свойство из списка доступных свойств
        PartsDataGrid.Items.Refresh();
    }

    private void AddCustomIPropertyColumn(string propertyName)
    {
        // Проверяем, существует ли уже колонка с таким именем или заголовком
        if (PartsDataGrid.Columns.Any(c => c.Header as string == propertyName))
        {
            MessageBox.Show($"Столбец с именем '{propertyName}' уже существует.", "Предупреждение", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        // Создаем новую колонку
        var column = new DataGridTextColumn
        {
            Header = propertyName,
            Binding = new Binding($"CustomProperties[{propertyName}]")
            {
                Mode = BindingMode.OneWay // Устанавливаем режим привязки
            },
            Width = DataGridLength.Auto
        };

        // Применяем стиль CenteredCellStyle
        column.ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style;

        PartsDataGrid.Columns.Add(column);
    }

    private void EditIProperty_Click(object sender, RoutedEventArgs e)
    {
        var selectedItem = PartsDataGrid.SelectedItem as PartData;
        if (selectedItem == null || PartsDataGrid.SelectedItems.Count != 1)
        {
            MessageBox.Show("Выберите одну строку для редактирования iProperty.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var editDialog = new EditIPropertyDialog(selectedItem.PartNumber, selectedItem.Description);
        if (editDialog.ShowDialog() == true)
        {
            // Открываем документ детали
            var partDoc = OpenPartDocument(selectedItem.PartNumber);
            if (partDoc != null)
            {
                // Обновляем iProperty в файле детали
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number", editDialog.PartNumber);
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description", editDialog.Description);

                // Сохраняем изменения и закрываем документ
                partDoc.Save();
                // partDoc.Close(); // Можно закрыть документ, если он не нужен больше открыт

                // Обновляем свойства детали в таблице
                selectedItem.PartNumber = editDialog.PartNumber;
                selectedItem.Description = editDialog.Description;

                // Обновляем данные в таблице
                PartsDataGrid.Items.Refresh();
            }
            else
            {
                MessageBox.Show("Не удалось открыть документ детали для редактирования.", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void SetProperty(PropertySet propertySet, string propertyName, string value)
    {
        try
        {
            var property = propertySet[propertyName];
            property.Value = value;
        }
        catch (Exception)
        {
            MessageBox.Show($"Не удалось обновить свойство {propertyName}.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private string GetCustomIPropertyValue(string partNumber, string propertyName)
    {
        try
        {
            var partDoc = OpenPartDocument(partNumber);
            if (partDoc != null)
            {
                var propertySet = partDoc.PropertySets["Inventor User Defined Properties"];

                // Проверяем, существует ли свойство в данном PropertySet
                if (propertySet[propertyName] is Property property) return property.Value?.ToString() ?? string.Empty;
            }
        }
        catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
        {
            // Ловим ошибку и возвращаем пустую строку
            return string.Empty;
        }
        catch (Exception ex)
        {
            MessageBoxHelper.ShowErrorMessage($"Ошибка при получении свойства {propertyName}: {ex.Message}");
        }

        return string.Empty;
    }
}

public class PartData : INotifyPropertyChanged
{
    // Пользовательские iProperty
    private Dictionary<string, string> customProperties = new();
    private string description;
    private bool isModified; // Флаг для отслеживания изменений
    private bool isReadOnly = true;
    private int item;
    private string material;
    private string originalPartNumber; // Хранит оригинальный PartNumber
    private string partNumber;
    private int quantity;
    private Brush quantityColor;
    private string thickness;

    // Выражения (если они есть)
    public string PartNumberExpression { get; set; }
    public string DescriptionExpression { get; set; }

    public Dictionary<string, string> CustomProperties
    {
        get => customProperties;
        set
        {
            customProperties = value;
            OnPropertyChanged();
        }
    }

    // Выражения для пользовательских iProperty
    public Dictionary<string, string> CustomPropertyExpressions { get; set; } = new();

    // Порядковый номер элемента
    public int Item
    {
        get => item;
        set
        {
            item = value;
            OnPropertyChanged();
        }
    }

    // Основные свойства детали
    public string PartNumber
    {
        get => partNumber;
        set
        {
            if (partNumber != value)
            {
                partNumber = value;
                IsModified = true; // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }
    }

    public string OriginalPartNumber
    {
        get => originalPartNumber;
        set
        {
            originalPartNumber = value;
            OnPropertyChanged();
        }
    }

    public string ModelState { get; set; }

    public string Description
    {
        get => description;
        set
        {
            if (description != value)
            {
                description = value;
                IsModified = true; // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }
    }

    public string Material
    {
        get => material;
        set
        {
            material = value;
            IsModified = true; // Устанавливаем флаг изменений при изменении
            OnPropertyChanged();
        }
    }

    public string Thickness
    {
        get => thickness;
        set
        {
            thickness = value;
            IsModified = true; // Устанавливаем флаг изменений при изменении
            OnPropertyChanged();
        }
    }

    public int OriginalQuantity { get; set; }

    public int Quantity
    {
        get => quantity;
        set
        {
            quantity = value;
            IsModified = true; // Устанавливаем флаг изменений при изменении
            OnPropertyChanged();
        }
    }

    // Свойства для отображения изображений и цветов
    public BitmapImage Preview { get; set; }
    public BitmapImage DxfPreview { get; set; }
    public Brush FlatPatternColor { get; set; }
    public Brush ProcessingColor { get; set; } = Brushes.Transparent;

    public Brush QuantityColor
    {
        get => quantityColor;
        set
        {
            quantityColor = value;
            OnPropertyChanged();
        }
    }

    // Новые свойства для хранения значений iProperty
    public string Author { get; set; }
    public string Revision { get; set; }
    public string Project { get; set; }
    public string StockNumber { get; set; }

    // Методы для установки значений без модификации флага IsModified
    public void SetPartNumberWithoutModification(string value)
    {
        if (partNumber != value)
        {
            partNumber = value;
            OnPropertyChanged(nameof(PartNumber));
        }
    }

    public void SetDescriptionWithoutModification(string value)
    {
        if (description != value)
        {
            description = value;
            OnPropertyChanged(nameof(Description));
        }
    }

    public void SetCustomPropertyWithoutModification(string propertyName, string value)
    {
        if (CustomProperties.ContainsKey(propertyName))
        {
            if (CustomProperties[propertyName] != value)
            {
                CustomProperties[propertyName] = value;
                OnPropertyChanged(nameof(CustomProperties));
            }
        }
        else
        {
            CustomProperties.Add(propertyName, value);
            OnPropertyChanged(nameof(CustomProperties));
        }
    }

    // Свойство для контроля редактируемости
    public bool IsReadOnly
    {
        get => isReadOnly;
        set
        {
            if (isReadOnly != value)
            {
                isReadOnly = value;
                OnPropertyChanged();
            }
        }
    }

    // Флаг для отслеживания изменений
    public bool IsModified
    {
        get => isModified;
        set
        {
            if (isModified != value)
            {
                isModified = value;
                OnPropertyChanged();
            }
        }
    }

    // Флаг переопределения
    public bool IsOverridden { get; set; }

    // Реализация интерфейса INotifyPropertyChanged
    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }


// Методы для работы с пользовательскими свойствами
public void AddCustomProperty(string propertyName, string propertyValue, string expressionValue = null)
    {
        // Добавляем или обновляем значение кастомного свойства
        if (CustomProperties.ContainsKey(propertyName))
            CustomProperties[propertyName] = propertyValue;
        else
            CustomProperties.Add(propertyName, propertyValue);

        // Добавляем или обновляем выражение для кастомного свойства
        if (expressionValue != null)
        {
            if (CustomPropertyExpressions.ContainsKey(propertyName))
                CustomPropertyExpressions[propertyName] = expressionValue;
            else
                CustomPropertyExpressions.Add(propertyName, expressionValue);
        }
        else
        {
            // Если выражение не указано, инициализируем его пустым значением
            if (!CustomPropertyExpressions.ContainsKey(propertyName))
                CustomPropertyExpressions.Add(propertyName, string.Empty);
        }

        IsModified = true; // Устанавливаем флаг изменений
        OnPropertyChanged(nameof(CustomProperties));
        OnPropertyChanged(nameof(CustomPropertyExpressions));
    }


    public void RemoveCustomProperty(string propertyName)
    {
        if (customProperties.ContainsKey(propertyName))
        {
            customProperties.Remove(propertyName);
            IsModified = true; // Устанавливаем флаг изменений
            OnPropertyChanged(nameof(CustomProperties));
        }
    }
}

public class HeaderAdorner : Adorner
{
    private readonly VisualCollection _visuals;

    public HeaderAdorner(UIElement adornedElement, UIElement header)
        : base(adornedElement)
    {
        _visuals = new VisualCollection(this);
        Child = header;
        _visuals.Add(Child);

        IsHitTestVisible = false;
        Opacity = 0.8; // Полупрозрачность
    }

    public UIElement Child { get; }

    protected override int VisualChildrenCount => _visuals.Count;

    protected override Size ArrangeOverride(Size finalSize)
    {
        Child.Arrange(new Rect(finalSize));
        return finalSize;
    }

    protected override Visual GetVisualChild(int index)
    {
        return _visuals[index];
    }
}

public static class MessageBoxHelper
{
    public static void ShowInventorNotRunningError()
    {
        MessageBox.Show("Inventor не запущен. Пожалуйста, запустите Inventor и попробуйте снова.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public static void ShowNoDocumentOpenWarning()
    {
        MessageBox.Show("Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public static void ShowScanningCancelledInfo()
    {
        MessageBox.Show("Процесс сканирования был прерван.", "Информация", MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    public static void ShowExportCancelledInfo()
    {
        MessageBox.Show("Процесс экспорта был прерван.", "Информация", MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    public static void ShowNoFlatPatternWarning()
    {
        MessageBox.Show("Выбранные файлы не содержат разверток.", "Информация", MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    public static void ShowExportCompletedInfo(int processedCount, int skippedCount, string elapsedTime)
    {
        MessageBox.Show(
            $"Экспорт DXF завершен.\nВсего файлов обработано: {processedCount + skippedCount}\nПропущено (без разверток): {skippedCount}\nВсего экспортировано: {processedCount}\nВремя выполнения: {elapsedTime}",
            "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    public static void ShowErrorMessage(string message)
    {
        MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public static void ShowNotSheetMetalWarning()
    {
        MessageBox.Show("Активный документ не является листовым металлом.", "Ошибка", MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }
}

public class PresetIProperty
{
    public string InternalName { get; set; } // Внутреннее имя колонки (например, "#")
    public string DisplayName { get; set; } // Псевдоним для отображения в списке выбора (например, "#Нумерация")
    public string InventorPropertyName { get; set; } // Соответствующее имя свойства iProperty в Inventor

    public bool ShouldBeAddedOnInit { get; set; } =
        false; // Новый флаг для определения необходимости добавления при старте
}