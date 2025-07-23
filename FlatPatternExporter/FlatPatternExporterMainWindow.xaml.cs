using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
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
using Application = Inventor.Application;
using Binding = System.Windows.Data.Binding;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using DragEventArgs = System.Windows.DragEventArgs;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Image = System.Windows.Controls.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;

namespace FlatPatternExporter;

public partial class FlatPatternExporterMainWindow : Window
{
    private bool _hasMissingReferences = false;
    private AdornerLayer _adornerLayer;
    private HeaderAdorner _headerAdorner;
    private bool _isColumnDraggedOutside;
    private DataGridColumn _reorderingColumn;
    private string _actualSearchText = string.Empty; // Поле для хранения фактического текста поиска

    public readonly List<string> _customPropertiesList = new();
    private string _fixedFolderPath = string.Empty;
    private bool _isCancelled;
    private bool _isCtrlPressed;
    private bool _isScanning;
    private bool _excludeReferenceParts = true;
    private bool _excludePurchasedParts = true;
    private bool _includeLibraryComponents = false;
    private List<string> _libraryPaths = new List<string>();
    private int _itemCounter = 1; // Инициализация счетчика пунктов
    private Document? _lastScannedDocument;
    private readonly ObservableCollection<PartData> _partsData = new();
    private readonly CollectionViewSource _partsDataView;
    private readonly DispatcherTimer _searchDelayTimer;
    public Application _thisApplication;
    private List<PartData> _conflictingParts = new(); // Список конфликтующих файлов
    private Dictionary<string, List<(string PartNumber, string FileName)>> _conflictFileDetails = new();

    public FlatPatternExporterMainWindow()
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
        ExcludeReferencePartsCheckBox.Checked += UpdateExcludeReferencePartsState;
        ExcludeReferencePartsCheckBox.Unchecked += UpdateExcludeReferencePartsState;
        ExcludePurchasedPartsCheckBox.Checked += UpdateExcludePurchasedPartsState;
        ExcludePurchasedPartsCheckBox.Unchecked += UpdateExcludePurchasedPartsState;
        IncludeLibraryComponentsCheckBox.Checked += UpdateIncludeLibraryComponentsState;
        IncludeLibraryComponentsCheckBox.Unchecked += UpdateIncludeLibraryComponentsState;
        // Инициализация путей библиотек
        InitializeLibraryPaths();

        PartsDataGrid.DragOver += PartsDataGrid_DragOver;
        PartsDataGrid.DragLeave += PartsDataGrid_DragLeave;

        // Инициализация CollectionViewSource для фильтрации
        _partsDataView = new CollectionViewSource { Source = _partsData };
        _partsDataView.Filter += PartsData_Filter;

        // Устанавливаем ItemsSource для DataGrid
        PartsDataGrid.ItemsSource = _partsDataView.View;



        // Настраиваем таймер для задержки поиска
        _searchDelayTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1) // 1 секунда задержки
        };
        _searchDelayTimer.Tick += SearchDelayTimer_Tick;


        // Инициализируем настройки слоев
        LayerSettings = LayerSettingsHelper.InitializeLayerSettings();
        AvailableColors = LayerSettingsHelper.GetAvailableColors();
        LineTypes = LayerSettingsHelper.GetLineTypes();


        // Инициализация предустановленных колонок
        PresetIProperties = new ObservableCollection<PresetIProperty>
{
    new() { InternalName = "Имя файла", DisplayName = "Имя файла", InventorPropertyName = "FileName" },
    new() { InternalName = "Полное имя файла", DisplayName = "Полное имя файла", InventorPropertyName = "FullFileName" },
    new() { InternalName = "ID", DisplayName = "Нумерация", InventorPropertyName = "Item" },
    new() { InternalName = "Обозначение", DisplayName = "Обозначение", InventorPropertyName = "PartNumber" },
    new() { InternalName = "Наименование", DisplayName = "Наименование", InventorPropertyName = "Description" },
    new() { InternalName = "Состояние модели", DisplayName = "Состояние модели", InventorPropertyName = "ModelState" },
    new() { InternalName = "Материал", DisplayName = "Материал", InventorPropertyName = "Material" },
    new() { InternalName = "Толщина", DisplayName = "Толщина", InventorPropertyName = "Thickness" },
    new() { InternalName = "Кол.", DisplayName = "Количество", InventorPropertyName = "Quantity" },
    new() { InternalName = "Изобр. детали", DisplayName = "Изображение детали", InventorPropertyName = "Preview" },
    new() { InternalName = "Изобр. развертки", DisplayName = "Изображение развертки", InventorPropertyName = "DxfPreview" },
    new() { InternalName = "Автор", DisplayName = "Автор", InventorPropertyName = "Author" },
    new() { InternalName = "Ревизия", DisplayName = "Ревизия", InventorPropertyName = "Revision" },
    new() { InternalName = "Проект", DisplayName = "Проект", InventorPropertyName = "Project" },
    new() { InternalName = "Инвентарный номер", DisplayName = "Инвентарный номер", InventorPropertyName = "StockNumber" },
    new() { InternalName = "Название", DisplayName = "Название", InventorPropertyName = "Title" },
    new() { InternalName = "Тема", DisplayName = "Тема", InventorPropertyName = "Subject" },
    new() { InternalName = "Ключевые слова", DisplayName = "Ключевые слова", InventorPropertyName = "Keywords" },
    new() { InternalName = "Примечание", DisplayName = "Примечание", InventorPropertyName = "Comments" },
    new() { InternalName = "Категория", DisplayName = "Категория", InventorPropertyName = "Category" },
    new() { InternalName = "Менеджер", DisplayName = "Менеджер", InventorPropertyName = "Manager" },
    new() { InternalName = "Компания", DisplayName = "Компания", InventorPropertyName = "Company" },
    new() { InternalName = "Время создания", DisplayName = "Время создания", InventorPropertyName = "CreationTime" },
    new() { InternalName = "Сметчик", DisplayName = "Сметчик", InventorPropertyName = "CostCenter" },
    new() { InternalName = "Проверил", DisplayName = "Проверил", InventorPropertyName = "CheckedBy" },
    new() { InternalName = "Нормоконтроль", DisplayName = "Нормоконтроль", InventorPropertyName = "EngApprovedBy" },
    new() { InternalName = "Статус", DisplayName = "Статус", InventorPropertyName = "UserStatus" },
    new() { InternalName = "Веб-ссылка", DisplayName = "Веб-ссылка", InventorPropertyName = "CatalogWebLink" },
    new() { InternalName = "Поставщик", DisplayName = "Поставщик", InventorPropertyName = "Vendor" },
    new() { InternalName = "Утвердил", DisplayName = "Утвердил", InventorPropertyName = "MfgApprovedBy" },
    new() { InternalName = "Статус разработки", DisplayName = "Статус разработки", InventorPropertyName = "DesignStatus" },
    new() { InternalName = "Проектировщик", DisplayName = "Проектировщик", InventorPropertyName = "Designer" },
    new() { InternalName = "Инженер", DisplayName = "Инженер", InventorPropertyName = "Engineer" },
    new() { InternalName = "Нач. отдела", DisplayName = "Нач. отдела", InventorPropertyName = "Authority" },
    new() { InternalName = "Масса", DisplayName = "Масса", InventorPropertyName = "Mass" },
    new() { InternalName = "Площадь поверхности", DisplayName = "Площадь поверхности", InventorPropertyName = "SurfaceArea" },
    new() { InternalName = "Объем", DisplayName = "Объем", InventorPropertyName = "Volume" },
    new() { InternalName = "Правило ЛМ", DisplayName = "Правило листового металла", InventorPropertyName = "SheetMetalRule" },
    new() { InternalName = "Ширина развертки", DisplayName = "Ширина развертки", InventorPropertyName = "FlatPatternWidth" },
    new() { InternalName = "Длинна развертки", DisplayName = "Длинна развертки", InventorPropertyName = "FlatPatternLength" },
    new() { InternalName = "Площадь развертки", DisplayName = "Площадь развертки", InventorPropertyName = "FlatPatternArea" },
    new() { InternalName = "Отделка", DisplayName = "Отделка", InventorPropertyName = "Appearance" }
};

        // Устанавливаем DataContext для текущего окна
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
        var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        Console.WriteLine($"Raw Informational Version: {informationalVersion}");

        if (informationalVersion != null && informationalVersion.Contains("build:"))
        {
            var parts = informationalVersion.Split(new[] { "build:" }, StringSplitOptions.None);
            if (parts.Length > 1)
            {
                var version = parts[0].Trim();
                var shortCommitHash = parts[1].Split('+')[0].Trim(); // Берем только часть до '+'
                return $"{version} build:{shortCommitHash}";
            }
        }

        return informationalVersion ?? "Версия неизвестна";
    }

    /// <summary>
    /// Централизованный метод для управления интерфейсом Inventor
    /// </summary>
    /// <param name="disableInteraction">true - отключить взаимодействие с пользователем, false - включить</param>
    public void SetInventorUserInterfaceState(bool disableInteraction)
    {
        if (_thisApplication?.UserInterfaceManager != null)
        {
            _thisApplication.UserInterfaceManager.UserInteractionDisabled = disableInteraction;
        }
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

    private async Task<string> GetPropertyExpressionOrValueAsync(PartDocument partDoc, string propertyName)
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
                        return prop.Value?.ToString() ?? string.Empty;
                    }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при получении значения для {propertyName}: {ex.Message}");
            }

            return string.Empty;
        });
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
        // Очищаем текстовое поле
        SearchTextBox.Text = string.Empty;
        ClearSearchButton.Visibility = Visibility.Collapsed;
        SearchTextBox.Focus(); // Устанавливаем фокус на поле поиска
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _actualSearchText = SearchTextBox.Text.Trim().ToLower();

        // Показываем или скрываем кнопку очистки в зависимости от наличия текста
        ClearSearchButton.Visibility =
            string.IsNullOrEmpty(_actualSearchText) ? Visibility.Collapsed : Visibility.Visible;

        // Перезапуск таймера при изменении текста
        _searchDelayTimer.Stop();
        _searchDelayTimer.Start();
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
            _thisApplication = (Application)MarshalCore.GetActiveObject("Inventor.Application");
            if (_thisApplication != null)
            {
                InitializeProjectAndLibraryData();
            }
        }
        catch (COMException)
        {
            MessageBox.Show(
                "Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Произошла ошибка при подключении к Inventor: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
    private void InitializeProjectAndLibraryData()
    {
        if (_thisApplication != null)
        {
            try
            {
                SetProjectFolderInfo();
                InitializeLibraryPaths();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при инициализации данных проекта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        if (_thisApplication == null)
        {
            InitializeInventor();
            if (_thisApplication == null)
            {
                MessageBox.Show("Inventor не запущен. Пожалуйста, запустите Inventor и попробуйте снова.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        // Проверка на наличие активного документа
        Document? doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBox.Show("Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetInventorUserInterfaceState(true);

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
        _hasMissingReferences = false; // Сброс флага потерянных ссылок

        var stopwatch = Stopwatch.StartNew();

        await Task.Delay(100);

        var partCount = 0;
        var documentType = doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "Сборка" : "Деталь";
        var partNumber = string.Empty;
        var description = string.Empty;
        var modelStateInfo = string.Empty;

        _partsData.Clear();
        _itemCounter = 1;

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
            (partNumber, description) = GetDocumentProperties((Document)asmDoc);
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
                (partNumber, description) = GetDocumentProperties((Document)partDoc);

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

        SetInventorUserInterfaceState(false);

        if (_isCancelled)
        {
            MessageBox.Show("Процесс сканирования был прерван.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        else if (_hasMissingReferences)
        {
            MessageBox.Show("В сборке обнаружены компоненты с потерянными ссылками. Некоторые данные могли быть пропущены.",
                "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
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
        var tasks = new List<Task>();

        foreach (BOMRow row in bomView.BOMRows)
        {
            foreach (ComponentDefinition componentDefinition in row.ComponentDefinitions)
            {
                tasks.Add(Task.Run(() =>
                {
                    if (_isCancelled) return;

                    try
                    {
                        var document = componentDefinition.Document as Document;
                        if (document == null) return;

                        if (ShouldExcludeComponent(row.BOMStructure, document.FullFileName)) return;

                        if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                        {
                            var partDocument = document as PartDocument;

                            if (partDocument != null && partDocument.ComponentDefinition is SheetMetalComponentDefinition)
                            {
                                var partNumber = GetProperty(partDocument.PropertySets["Design Tracking Properties"], "Part Number");
                                var fileName = partDocument.FullFileName;

                                partNumbers.AddOrUpdate(partNumber,
                                    new List<(string PartNumber, string FileName)> { (partNumber, fileName) },
                                    (key, existingList) =>
                                    {
                                        if (!existingList.Any(p => p.FileName == fileName))
                                        {
                                            existingList.Add((partNumber, fileName));
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
                    catch (Exception ex)
                    {
                        // Логирование ошибки
                        Debug.WriteLine($"Ошибка при проверке конфликтов: {ex.Message}");
                    }
                }));
            }
        }

        await Task.WhenAll(tasks);
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
        if (_thisApplication == null) return;

        try
        {
            var activeProject = _thisApplication.DesignProjectManager.ActiveDesignProject;
            var projectName = activeProject.Name;
            var projectWorkspacePath = activeProject.WorkspacePath;
            ProjectName.Text = $"Проект: {projectName}";
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

    private (string partNumber, string description) GetDocumentProperties(Document document)
    {
        var partNumber = GetProperty(document.PropertySets["Design Tracking Properties"], "Part Number");
        var description = GetProperty(document.PropertySets["Design Tracking Properties"], "Description");
        return (partNumber, description);
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
}

public class PartData : INotifyPropertyChanged
{
    // Пользовательские iProperty
    private Dictionary<string, string> customProperties = new();
    private string description;
    private int item;
    private string material;
    private string partNumber;
    private int quantity;
    private Brush quantityColor;
    private string thickness;
    public string FileName { get; set; }
    public string FullFileName { get; set; }
    public string Author { get; set; }
    public string Revision { get; set; }
    public string Project { get; set; }
    public string StockNumber { get; set; }
    public string Title { get; set; }
    public string Subject { get; set; }
    public string Keywords { get; set; }
    public string Comments { get; set; }
    public string Category { get; set; }
    public string Manager { get; set; }
    public string Company { get; set; }
    public string CreationTime { get; set; }
    public string CostCenter { get; set; }
    public string CheckedBy { get; set; }
    public string EngApprovedBy { get; set; }
    public string UserStatus { get; set; }
    public string CatalogWebLink { get; set; }
    public string Vendor { get; set; }
    public string MfgApprovedBy { get; set; }
    public string DesignStatus { get; set; }
    public string Designer { get; set; }
    public string Engineer { get; set; }
    public string Authority { get; set; }
    public string Mass { get; set; }
    public string SurfaceArea { get; set; }
    public string Volume { get; set; }
    public string SheetMetalRule { get; set; }
    public string FlatPatternWidth { get; set; }
    public string FlatPatternLength { get; set; }
    public string FlatPatternArea { get; set; }
    public string Appearance { get; set; }


    public Dictionary<string, string> CustomProperties
    {
        get => customProperties;
        set
        {
            customProperties = value;
            OnPropertyChanged();
        }
    }


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
                OnPropertyChanged();
            }
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
            OnPropertyChanged();
        }
    }

    public string Thickness
    {
        get => thickness;
        set
        {
            thickness = value;
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



    // Флаг переопределения
    public bool IsOverridden { get; set; }

    // Реализация интерфейса INotifyPropertyChanged
    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // Методы для работы с пользовательскими свойствами
    public void AddCustomProperty(string propertyName, string propertyValue)
    {
        if (CustomProperties.ContainsKey(propertyName))
            CustomProperties[propertyName] = propertyValue;
        else
            CustomProperties.Add(propertyName, propertyValue);

        OnPropertyChanged(nameof(CustomProperties));
    }
    public void RemoveCustomProperty(string propertyName)
    {
        if (customProperties.ContainsKey(propertyName))
        {
            customProperties.Remove(propertyName);
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


public class PresetIProperty
{
    public string InternalName { get; set; } // Внутреннее имя колонки (например, "#")
    public string DisplayName { get; set; } // Псевдоним для отображения в списке выбора (например, "#Нумерация")
    public string InventorPropertyName { get; set; } // Соответствующее имя свойства iProperty в Inventor

    public bool ShouldBeAddedOnInit { get; set; } =
        false; // Новый флаг для определения необходимости добавления при старте
}