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
using DragEventArgs = System.Windows.DragEventArgs;
using Image = System.Windows.Controls.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;

namespace FlatPatternExporter;

public partial class FlatPatternExporterMainWindow : Window, INotifyPropertyChanged
{
    private bool _hasMissingReferences = false;
    private AdornerLayer? _adornerLayer;
    private HeaderAdorner? _headerAdorner;
    private bool _isColumnDraggedOutside;
    private DataGridColumn? _reorderingColumn;
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
    public Application? _thisApplication;
    private List<PartData> _conflictingParts = new(); // Список конфликтующих файлов
    private Dictionary<string, List<(string PartNumber, string FileName)>> _conflictFileDetails = new();
    private readonly ConcurrentDictionary<string, List<(string PartNumber, string FileName)>> _partNumberTracker = new(); // Для отслеживания конфликтов во время сканирования

    public FlatPatternExporterMainWindow()
    {
        InitializeComponent();
        InitializeInventor();
        InitializeProjectFolder(); // Инициализация папки проекта при запуске
        PartsDataGrid.ItemsSource = _partsData;
        MultiplierTextBox.IsEnabled = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        UpdateFileCountLabel(0);
        ClearButton.IsEnabled = false;
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


        // Инициализация предустановленных колонок с указанием категорий
        PresetIProperties = new ObservableCollection<PresetIProperty>
{
    // КАТЕГОРИЯ 1: Системные свойства приложения
    new() { InternalName = "ID", DisplayName = "Нумерация", InventorPropertyName = "Item", Category = "Системные" },
    
    // КАТЕГОРИЯ 2: Свойства документа (не iProperty)
    new() { InternalName = "Имя файла", DisplayName = "Имя файла", InventorPropertyName = "FileName", Category = "Документ" },
    new() { InternalName = "Полное имя файла", DisplayName = "Полное имя файла", InventorPropertyName = "FullFileName", Category = "Документ" },
    new() { InternalName = "Состояние модели", DisplayName = "Состояние модели", InventorPropertyName = "ModelState", Category = "Документ" },
    new() { InternalName = "Толщина", DisplayName = "Толщина", InventorPropertyName = "Thickness", Category = "Документ" },
    new() { InternalName = "Изобр. детали", DisplayName = "Изображение детали", InventorPropertyName = "Preview", Category = "Документ" },
    
    // КАТЕГОРИЯ 3: Основные iProperty
    new() { InternalName = "Обозначение", DisplayName = "Обозначение", InventorPropertyName = "PartNumber", Category = "Основные iProperty" },
    new() { InternalName = "Наименование", DisplayName = "Наименование", InventorPropertyName = "Description", Category = "Основные iProperty" },
    new() { InternalName = "Материал", DisplayName = "Материал", InventorPropertyName = "Material", Category = "Основные iProperty" },
    
    // КАТЕГОРИЯ 4: Расширенные iProperty - Summary Information
    new() { InternalName = "Автор", DisplayName = "Автор", InventorPropertyName = "Author", Category = "Summary Information" },
    new() { InternalName = "Ревизия", DisplayName = "Ревизия", InventorPropertyName = "Revision", Category = "Summary Information" },
    new() { InternalName = "Название", DisplayName = "Название", InventorPropertyName = "Title", Category = "Summary Information" },
    new() { InternalName = "Тема", DisplayName = "Тема", InventorPropertyName = "Subject", Category = "Summary Information" },
    new() { InternalName = "Ключевые слова", DisplayName = "Ключевые слова", InventorPropertyName = "Keywords", Category = "Summary Information" },
    new() { InternalName = "Примечание", DisplayName = "Примечание", InventorPropertyName = "Comments", Category = "Summary Information" },
    
    // КАТЕГОРИЯ 4: Расширенные iProperty - Document Summary Information
    new() { InternalName = "Категория", DisplayName = "Категория", InventorPropertyName = "Category", Category = "Document Summary Information" },
    new() { InternalName = "Менеджер", DisplayName = "Менеджер", InventorPropertyName = "Manager", Category = "Document Summary Information" },
    new() { InternalName = "Компания", DisplayName = "Компания", InventorPropertyName = "Company", Category = "Document Summary Information" },
    
    // КАТЕГОРИЯ 4: Расширенные iProperty - Design Tracking Properties
    new() { InternalName = "Проект", DisplayName = "Проект", InventorPropertyName = "Project", Category = "Design Tracking Properties" },
    new() { InternalName = "Инвентарный номер", DisplayName = "Инвентарный номер", InventorPropertyName = "StockNumber", Category = "Design Tracking Properties" },
    new() { InternalName = "Время создания", DisplayName = "Время создания", InventorPropertyName = "CreationTime", Category = "Design Tracking Properties" },
    new() { InternalName = "Сметчик", DisplayName = "Сметчик", InventorPropertyName = "CostCenter", Category = "Design Tracking Properties" },
    new() { InternalName = "Проверил", DisplayName = "Проверил", InventorPropertyName = "CheckedBy", Category = "Design Tracking Properties" },
    new() { InternalName = "Нормоконтроль", DisplayName = "Нормоконтроль", InventorPropertyName = "EngApprovedBy", Category = "Design Tracking Properties" },
    new() { InternalName = "Статус", DisplayName = "Статус", InventorPropertyName = "UserStatus", Category = "Design Tracking Properties" },
    new() { InternalName = "Веб-ссылка", DisplayName = "Веб-ссылка", InventorPropertyName = "CatalogWebLink", Category = "Design Tracking Properties" },
    new() { InternalName = "Поставщик", DisplayName = "Поставщик", InventorPropertyName = "Vendor", Category = "Design Tracking Properties" },
    new() { InternalName = "Утвердил", DisplayName = "Утвердил", InventorPropertyName = "MfgApprovedBy", Category = "Design Tracking Properties" },
    new() { InternalName = "Статус разработки", DisplayName = "Статус разработки", InventorPropertyName = "DesignStatus", Category = "Design Tracking Properties" },
    new() { InternalName = "Проектировщик", DisplayName = "Проектировщик", InventorPropertyName = "Designer", Category = "Design Tracking Properties" },
    new() { InternalName = "Инженер", DisplayName = "Инженер", InventorPropertyName = "Engineer", Category = "Design Tracking Properties" },
    new() { InternalName = "Нач. отдела", DisplayName = "Нач. отдела", InventorPropertyName = "Authority", Category = "Design Tracking Properties" },
    new() { InternalName = "Масса", DisplayName = "Масса", InventorPropertyName = "Mass", Category = "Design Tracking Properties" },
    new() { InternalName = "Площадь поверхности", DisplayName = "Площадь поверхности", InventorPropertyName = "SurfaceArea", Category = "Design Tracking Properties" },
    new() { InternalName = "Объем", DisplayName = "Объем", InventorPropertyName = "Volume", Category = "Design Tracking Properties" },
    new() { InternalName = "Правило ЛМ", DisplayName = "Правило листового металла", InventorPropertyName = "SheetMetalRule", Category = "Design Tracking Properties" },
    new() { InternalName = "Ширина развертки", DisplayName = "Ширина развертки", InventorPropertyName = "FlatPatternWidth", Category = "Design Tracking Properties" },
    new() { InternalName = "Длинна развертки", DisplayName = "Длинна развертки", InventorPropertyName = "FlatPatternLength", Category = "Design Tracking Properties" },
    new() { InternalName = "Площадь развертки", DisplayName = "Площадь развертки", InventorPropertyName = "FlatPatternArea", Category = "Design Tracking Properties" },
    new() { InternalName = "Отделка", DisplayName = "Отделка", InventorPropertyName = "Appearance", Category = "Design Tracking Properties" },
    
    // КАТЕГОРИЯ 6: Свойства количества и состояния
    new() { InternalName = "Кол.", DisplayName = "Количество", InventorPropertyName = "Quantity", Category = "Количество" },
    
    // КАТЕГОРИЯ 7: Свойства состояния обработки
    new() { InternalName = "Изобр. развертки", DisplayName = "Изображение развертки", InventorPropertyName = "DxfPreview", Category = "Обработка" }
};

        // Устанавливаем DataContext для текущего окна
        DataContext = this;

        // Инициализация DataGrid с предустановленными колонками
        InitializePresetColumns();

        // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
        KeyDown += MainWindow_KeyDown;
        KeyUp += MainWindow_KeyUp;

        VersionTextBlock.Text = "Версия программы: " + GetVersion();
        ProgressLabel.Text = "Статус: Документ не выбран";
    }

    public ObservableCollection<LayerSetting> LayerSettings { get; set; }
    public ObservableCollection<string> AvailableColors { get; set; }
    public ObservableCollection<string> LineTypes { get; set; }
    public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // Свойство для привязки состояния модели к UI (устанавливается однократно при сканировании)
    private bool _isPrimaryModelState = true;
    public bool IsPrimaryModelState
    {
        get => _isPrimaryModelState;
        set
        {
            if (_isPrimaryModelState != value)
            {
                _isPrimaryModelState = value;
                OnPropertyChanged();
            }
        }
    }

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
    private void SearchDelayTimer_Tick(object? sender, EventArgs e)
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
                          (partData.Thickness.ToString("F1", System.Globalization.CultureInfo.InvariantCulture).ToLower().Contains(_actualSearchText) == true ||
                           partData.Thickness.ToString("F1", System.Globalization.CultureInfo.CurrentCulture).ToLower().Contains(_actualSearchText) == true ||
                           partData.Thickness.ToString("F0").ToLower().Contains(_actualSearchText) == true) ||
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
        if (iProperty.InternalName == "Изобр. детали" || iProperty.InternalName == "Изобр. развертки")
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
            (column as DataGridTemplateColumn)?.CellTemplate?.VisualTree?.SetBinding(Image.SourceProperty,
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

            // Стили теперь назначаются автоматически через неявный стиль DataGridCell в XAML

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

        // Дозаполняем данные для новой колонки (асинхронно)
        _ = FillPropertyDataAsync(iProperty.InventorPropertyName);
    }

    private void PartsDataGrid_ColumnReordering(object? sender, DataGridColumnReorderingEventArgs e)
    {
        _reorderingColumn = e.Column;
        _isColumnDraggedOutside = false;

        // Создаем фантомный заголовок используя стиль из XAML
        var header = new TextBlock
        {
            Text = _reorderingColumn.Header.ToString(),
            Width = _reorderingColumn.ActualWidth,
            Style = (Style)FindResource("PhantomColumnHeaderStyle")
        };

        // Создаем и добавляем Adorner
        _adornerLayer = AdornerLayer.GetAdornerLayer(PartsDataGrid);
        _headerAdorner = new HeaderAdorner(PartsDataGrid, header);
        _adornerLayer.Add(_headerAdorner);

        // Изначально размещаем Adorner там, где находится заголовок
        UpdateAdornerPosition(e);
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
            if (color == Brushes.Red)
            {
                textBlock.Text = $"❎ {_reorderingColumn?.Header}";
                textBlock.Style = (Style)FindResource("PhantomColumnHeaderDeleteStyle");
            }
            else
            {
                textBlock.Text = _reorderingColumn?.Header?.ToString() ?? string.Empty;
                textBlock.Style = (Style)FindResource("PhantomColumnHeaderStyle");
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
            if (string.IsNullOrEmpty(columnName)) return;

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

    private void PartsDataGrid_ColumnReordered(object? sender, DataGridColumnEventArgs e)
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

    private void UpdateAdornerPosition(DataGridColumnReorderingEventArgs? e)
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
        var textBox = sender as TextBox;
        if (textBox == null) return;
        var newText = textBox.Text.Insert(textBox.CaretIndex, e.Text);

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
            partData.IsMultiplied = multiplier > 1; // Устанавливаем флаг множителя
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


        if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject &&
            doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
        {
            MessageBox.Show("Откройте сборку или деталь для сканирования.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        ResetProgressBar();
        ProgressBar.IsIndeterminate = false;
        ProgressBar.Minimum = 0;
        ProgressBar.Maximum = 100;
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Статус: Подготовка к сканированию...";
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
        _partNumberTracker.Clear(); // Очищаем трекер конфликтов

        // Прогресс для сканирования структуры сборки
        var scanProgress = new Progress<ScanProgress>(progress =>
        {
            ProgressBar.Value = progress.TotalItems > 0 ? (double)progress.ProcessedItems / progress.TotalItems * 100 : 0;
            ProgressLabel.Text = $"Статус: {progress.CurrentOperation} - {progress.CurrentItem}";
        });

        // Прогресс для добавления готовых деталей в таблицу
        var partProgress = new Progress<PartData>(partData =>
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
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, scanProgress));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, scanProgress));
            partCount = sheetMetalParts.Count;
            (partNumber, description) = GetDocumentProperties((Document)asmDoc);
            modelStateInfo = asmDoc.ComponentDefinition.BOM.BOMViews[1].ModelStateMemberName;

            // Переключаем прогресс на обработку деталей
            await Dispatcher.InvokeAsync(() =>
            {
                ProgressBar.Value = 0;
                ProgressLabel.Text = "Статус: Обработка деталей...";
            });

            var itemCounter = 1;
            var totalParts = sheetMetalParts.Count;
            var processedParts = 0;
            
            await Task.Run(async () =>
            {
                foreach (var part in sheetMetalParts)
                {
                    if (_isCancelled) break;

                    var partData = await GetPartDataAsync(part.Key, part.Value, asmDoc.ComponentDefinition.BOM,
                        itemCounter++);
                    if (partData != null)
                    {
                        ((IProgress<PartData>)partProgress).Report(partData);
                    }
                    
                    // Обновляем прогресс обработки деталей
                    processedParts++;
                    await Dispatcher.InvokeAsync(() =>
                    {
                        ProgressBar.Value = totalParts > 0 ? (double)processedParts / totalParts * 100 : 0;
                        ProgressLabel.Text = $"Статус: Обработка деталей - {processedParts} из {totalParts}";
                    });
                    
                    await Task.Delay(10);
                }
            });

            // Вызов анализа конфликтов после завершения сканирования
            await AnalyzePartNumberConflictsAsync(stopwatch);
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                partCount = 1;
                (partNumber, description) = GetDocumentProperties((Document)partDoc);

                ProgressLabel.Text = "Статус: Обработка детали...";
                var partData = await GetPartDataAsync(partNumber, 1, null, 1, partDoc);
                if (partData != null)
                {
                    ((IProgress<PartData>)partProgress).Report(partData);
                    ProgressBar.Value = 100;
                    await Task.Delay(10);
                }
            }
        }

        // Останавливаем секундомер, если он еще не остановлен
        if (stopwatch.IsRunning)
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
        var caretIndex = textBox?.CaretIndex ?? 0;

        if (textBox != null)
        {
            foreach (var ch in forbiddenChars)
                if (textBox.Text.Contains(ch))
                {
                    textBox.Text = textBox.Text.Replace(ch.ToString(), string.Empty);
                    caretIndex--; // Уменьшаем позицию курсора на 1
                }
        }

        // Возвращаем курсор на правильное место после удаления символов
        if (textBox != null)
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
        var mgr = new PropertyManager((Document)document);
        var partNumber = mgr.GetMappedProperty("PartNumber");
        var description = mgr.GetMappedProperty("Description");
        return (partNumber, description);
    }

    private void AddPartToConflictTracker(string partNumber, string fileName)
    {
        _partNumberTracker.AddOrUpdate(partNumber,
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

    private async Task AnalyzePartNumberConflictsAsync(Stopwatch stopwatch)
    {
        _conflictingParts.Clear(); // Очищаем список конфликтов перед началом проверки

        // Оставляем только те записи, у которых больше одного файла (т.е. есть конфликт)
        var conflictingPartNumbers = _partNumberTracker.Where(p => p.Value.Count > 1).ToDictionary(p => p.Key, p => p.Value);

        if (conflictingPartNumbers.Any())
        {
            _conflictingParts.AddRange(conflictingPartNumbers.SelectMany(entry => entry.Value.Select(v => new PartData { PartNumber = v.PartNumber })));

            // Останавливаем секундомер перед показом MessageBox
            stopwatch.Stop();

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


}

public class PartData : INotifyPropertyChanged
{
    // === КАТЕГОРИЯ 1: СИСТЕМНЫЕ СВОЙСТВА ПРИЛОЖЕНИЯ ===
    // Свойства, используемые только внутри приложения для управления таблицей
    private int item;

    // === КАТЕГОРИЯ 2: СВОЙСТВА ДОКУМЕНТА (НЕ IPROPERTY) ===
    // Базовые свойства файла документа, считываемые напрямую из API Inventor
    public string FileName { get; set; } = string.Empty; // Path.GetFileNameWithoutExtension()
    public string FullFileName { get; set; } = string.Empty; // partDoc.FullFileName
    public string ModelState { get; set; } = string.Empty; // partDoc.ModelStateName
    private double thickness = 0.0; // SheetMetalComponentDefinition.Thickness
    public BitmapImage? Preview { get; set; } // apprenticeDoc.Thumbnail
    public bool HasFlatPattern { get; set; } = false; // smCompDef.HasFlatPattern

    // === КАТЕГОРИЯ 3: ОСНОВНЫЕ IPROPERTIES ===
    // Ключевые iProperty из "Design Tracking Properties", всегда необходимые
    private string partNumber = string.Empty;
    private string description = string.Empty;
    private string material = string.Empty;

    // === КАТЕГОРИЯ 4: РАСШИРЕННЫЕ IPROPERTIES ===
    // Опциональные iProperty из различных наборов свойств
    public string Author { get; set; } = string.Empty;
    public string Revision { get; set; } = string.Empty;
    public string Project { get; set; } = string.Empty;
    public string StockNumber { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Keywords { get; set; } = string.Empty;
    public string Comments { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Manager { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;
    public string CreationTime { get; set; } = string.Empty;
    public string CostCenter { get; set; } = string.Empty;
    public string CheckedBy { get; set; } = string.Empty;
    public string EngApprovedBy { get; set; } = string.Empty;
    public string UserStatus { get; set; } = string.Empty;
    public string CatalogWebLink { get; set; } = string.Empty;
    public string Vendor { get; set; } = string.Empty;
    public string MfgApprovedBy { get; set; } = string.Empty;
    public string DesignStatus { get; set; } = string.Empty;
    public string Designer { get; set; } = string.Empty;
    public string Engineer { get; set; } = string.Empty;
    public string Authority { get; set; } = string.Empty;
    public string Mass { get; set; } = string.Empty;
    public string SurfaceArea { get; set; } = string.Empty;
    public string Volume { get; set; } = string.Empty;
    public string SheetMetalRule { get; set; } = string.Empty;
    public string FlatPatternWidth { get; set; } = string.Empty;
    public string FlatPatternLength { get; set; } = string.Empty;
    public string FlatPatternArea { get; set; } = string.Empty;
    public string Appearance { get; set; } = string.Empty;

    // === КАТЕГОРИЯ 5: ПОЛЬЗОВАТЕЛЬСКИЕ IPROPERTIES ===
    // Динамически добавляемые пользователем свойства из "Inventor User Defined Properties"
    private Dictionary<string, string> customProperties = new();

    // === КАТЕГОРИЯ 6: СВОЙСТВА КОЛИЧЕСТВА И СОСТОЯНИЯ ===
    // Свойства для управления количеством и состоянием обработки
    private int quantity;
    public int OriginalQuantity { get; set; }
    public bool IsOverridden { get; set; }
    public bool IsMultiplied { get; set; }

    // === КАТЕГОРИЯ 7: СВОЙСТВА СОСТОЯНИЯ ОБРАБОТКИ ===
    // Свойства, устанавливаемые во время экспорта и обработки
    public BitmapImage? DxfPreview { get; set; } // Миниатюра развертки (генерируется)
    public Brush ProcessingColor { get; set; } = Brushes.Transparent; // Индикатор состояния экспорта

    // === СВОЙСТВА С УВЕДОМЛЕНИЯМИ ===

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

    // КАТЕГОРИЯ 3: Основные iProperty с уведомлениями
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

    public double Thickness
    {
        get => thickness;
        set
        {
            thickness = value;
            OnPropertyChanged();
        }
    }

    // КАТЕГОРИЯ 6: Свойства количества и состояния с уведомлениями
    public int Quantity
    {
        get => quantity;
        set
        {
            quantity = value;
            OnPropertyChanged();
        }
    }



    // Реализация интерфейса INotifyPropertyChanged
    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
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
    public string InternalName { get; set; } = string.Empty; // Внутреннее имя колонки (например, "#")
    public string DisplayName { get; set; } = string.Empty; // Псевдоним для отображения в списке выбора (например, "#Нумерация")
    public string InventorPropertyName { get; set; } = string.Empty; // Соответствующее имя свойства iProperty в Inventor
    public string Category { get; set; } = string.Empty; // Категория свойства для группировки
    public bool ShouldBeAddedOnInit { get; set; } = false; // Новый флаг для определения необходимости добавления при старте
}

// Структура для отслеживания прогресса сканирования
public class ScanProgress
{
    public int ProcessedItems { get; set; }
    public int TotalItems { get; set; }
    public string CurrentOperation { get; set; } = string.Empty;
    public string CurrentItem { get; set; } = string.Empty;
}

