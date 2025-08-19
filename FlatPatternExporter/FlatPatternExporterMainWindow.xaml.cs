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
using Binding = System.Windows.Data.Binding;
using DragEventArgs = System.Windows.DragEventArgs;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;

namespace FlatPatternExporter;

public enum ExportFolderType
{
    ChooseFolder = 0,      // Указать папку в процессе экспорта
    ComponentFolder = 1,   // Папка компонента
    PartFolder = 2,        // Рядом с деталью
    ProjectFolder = 3,     // Папка проекта
    FixedFolder = 4        // Фиксированная папка
}

public enum ProcessingMethod
{
    Traverse = 0,          // Перебор
    BOM = 1               // Спецификация
}

public class AcadVersionItem
{
    public string DisplayName { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    
    public override string ToString() => DisplayName;
}

public class SplineReplacementItem  
{
    public string DisplayName { get; set; } = string.Empty;
    public int Index { get; set; }
    
    public override string ToString() => DisplayName;
}

public partial class FlatPatternExporterMainWindow : Window, INotifyPropertyChanged
{
    // Inventor API
    public Inventor.Application? _thisApplication;
    private Document? _lastScannedDocument;
    
    // Сервисы
    private readonly TokenService _tokenService = new();
    
    // Данные и коллекции
    private readonly ObservableCollection<PartData> _partsData = new();
    private readonly CollectionViewSource _partsDataView;
    public readonly List<string> _customPropertiesList = new();
    public List<string> CustomPropertiesList => _customPropertiesList;
    private List<PartData> _conflictingParts = new();
    private Dictionary<string, List<PartConflictInfo>> _conflictFileDetails = new();
    private readonly ConcurrentDictionary<string, List<PartConflictInfo>> _partNumberTracker = new();
    
    // Состояние процессов
    private bool _isScanning;
    private bool _isExporting;
    private CancellationTokenSource? _operationCts;
    private bool _hasMissingReferences = false;
    private int _itemCounter = 1;
    private bool _isInitializing = true;
    
    // UI состояние
    private double _scanProgressValue;
    public double ScanProgressValue
    {
        get => _scanProgressValue;
        set
        {
            if (Math.Abs(_scanProgressValue - value) > 0.01)
            {
                _scanProgressValue = value;
                OnPropertyChanged();
            }
        }
    }
    
    private double _exportProgressValue;
    public double ExportProgressValue
    {
        get => _exportProgressValue;
        set
        {
            if (Math.Abs(_exportProgressValue - value) > 0.01)
            {
                _exportProgressValue = value;
                OnPropertyChanged();
            }
        }
    }
    private bool _isCtrlPressed;
    private string _actualSearchText = string.Empty;
    private readonly DispatcherTimer _searchDelayTimer;
    
    // DataGrid управление колонками
    private AdornerLayer? _adornerLayer;
    private HeaderAdorner? _headerAdorner;
    private bool _isColumnDraggedOutside;
    private DataGridColumn? _reorderingColumn;
    
    // Настройки фильтрации деталей
    private bool _excludeReferenceParts = true;
    private bool _excludePurchasedParts = true;
    private bool _excludePhantomParts = true;
    private bool _includeLibraryComponents = false;
    
    // Настройки организации файлов
    private bool _organizeByMaterial = false;
    private bool _organizeByThickness = false;
    
    // Настройки конструктора имени файла
    private bool _enableFileNameConstructor = false;
    
    // Настройки оптимизации
    private bool _optimizeDxf = false;
    
    // Настройки экспорта DXF
    private bool _enableSplineReplacement = false;
    private int _selectedSplineReplacementIndex = 0;
    private int _selectedAcadVersionIndex = 5; // 2000 по умолчанию
    private bool _mergeProfilesIntoPolyline = true;
    private bool _rebaseGeometry = true;
    private bool _trimCenterlines = false;
    
    // Настройки папок и путей
    private ExportFolderType _selectedExportFolder = ExportFolderType.ChooseFolder;
    private bool _enableSubfolder = false;
    private string _fixedFolderPath = string.Empty;
    public string FixedFolderPath 
    { 
        get => _fixedFolderPath; 
        set 
        {
            _fixedFolderPath = value;
            FixedFolderPathTextBlock.Text = value.Length > 55 ? $"... {value[^55..]}" : value;
        }
    }
    private List<string> _libraryPaths = new List<string>();
    
    // Метод обработки
    private ProcessingMethod _selectedProcessingMethod = ProcessingMethod.BOM;
    
    // Модель состояния
    private bool _isPrimaryModelState = true;
    
    // Словарь колонок с шаблонами (все типы)
    private static readonly Dictionary<string, string> ColumnTemplates = new()
    {
        { "Обр.", "ProcessingStatusTemplate" },
        { "Изобр. детали", "PartImageTemplate" },
        { "Изобр. развертки", "DxfImageTemplate" },
        { "Обозначение", "PartNumberWithIndicatorTemplate" },
        { "Кол.", "QuantityTemplate" }
    };

    public FlatPatternExporterMainWindow()
    {
        InitializeComponent();
        InitializeInventor();
        InitializeProjectFolder(); // Инициализация папки проекта при запуске
        PartsDataGrid.ItemsSource = _partsData;
        PartsDataGrid.PreviewMouseMove += PartsDataGrid_PreviewMouseMove;
        PartsDataGrid.PreviewMouseLeftButtonUp += PartsDataGrid_PreviewMouseLeftButtonUp;
        PartsDataGrid.ColumnReordering += PartsDataGrid_ColumnReordering;
        PartsDataGrid.ColumnReordered += PartsDataGrid_ColumnReordered;
        
        // Обновляем видимость оверлея при изменении коллекции колонок
        ((System.Collections.Specialized.INotifyCollectionChanged)PartsDataGrid.Columns).CollectionChanged += (s, e) => UpdateNoColumnsOverlayVisibility();
        
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
        
        // Инициализируем коллекции для ComboBox
        InitializeAcadVersions();
        InitializeSplineReplacementTypes();
        InitializeAvailableTokens();


        // Инициализация предустановленных колонок на основе PropertyMapping
        PresetIProperties = PropertyManager.GetPresetProperties();

        // Устанавливаем DataContext для текущего окна
        DataContext = this;

        // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
        KeyDown += MainWindow_KeyDown;
        KeyUp += MainWindow_KeyUp;

        SetUIState(UIState.Initial);
        
        // Загружаем настройки при запуске
        LoadSettings();
        
        // Инициализируем видимость оверлея
        UpdateNoColumnsOverlayVisibility();
        
        // Инициализируем данные в TokenService
        _tokenService.UpdatePartsData(_partsData);
        
        // Завершаем инициализацию
        _isInitializing = false;
    }

    private void LoadSettings()
    {
        try
        {
            var settings = SettingsManager.LoadSettings();
            SettingsManager.ApplySettingsToMainWindow(settings, this);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при загрузке настроек: {ex.Message}", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SaveSettings()
    {
        try
        {
            var settings = SettingsManager.CreateSettingsFromMainWindow(this);
            SettingsManager.SaveSettings(settings);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при сохранении настроек: {ex.Message}", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }


    public ObservableCollection<LayerSetting> LayerSettings { get; set; }
    public ObservableCollection<string> AvailableColors { get; set; }
    public ObservableCollection<string> LineTypes { get; set; }
    public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }
    public ObservableCollection<AcadVersionItem> AcadVersions { get; set; } = new();
    public ObservableCollection<SplineReplacementItem> SplineReplacementTypes { get; set; } = new();
    public ObservableCollection<string> AvailableTokens { get; set; } = new();
    public ObservableCollection<TemplatePreset> TemplatePresets { get; set; } = new();

    private TemplatePreset? _selectedTemplatePreset;
    public TemplatePreset? SelectedTemplatePreset
    {
        get => _selectedTemplatePreset;
        set
        {
            _selectedTemplatePreset = value;
            OnPropertyChanged();
        }
    }

    // Публичные свойства для привязки данных CheckBox
    public bool ExcludeReferenceParts
    {
        get => _excludeReferenceParts;
        set
        {
            if (_excludeReferenceParts != value)
            {
                _excludeReferenceParts = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ExcludePurchasedParts
    {
        get => _excludePurchasedParts;
        set
        {
            if (_excludePurchasedParts != value)
            {
                _excludePurchasedParts = value;
                OnPropertyChanged();
            }
        }
    }

    public bool ExcludePhantomParts
    {
        get => _excludePhantomParts;
        set
        {
            if (_excludePhantomParts != value)
            {
                _excludePhantomParts = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IncludeLibraryComponents
    {
        get => _includeLibraryComponents;
        set
        {
            if (_includeLibraryComponents != value)
            {
                _includeLibraryComponents = value;
                OnPropertyChanged();
            }
        }
    }

    public bool OrganizeByMaterial
    {
        get => _organizeByMaterial;
        set
        {
            if (_organizeByMaterial != value)
            {
                _organizeByMaterial = value;
                OnPropertyChanged();
            }
        }
    }

    public bool OrganizeByThickness
    {
        get => _organizeByThickness;
        set
        {
            if (_organizeByThickness != value)
            {
                _organizeByThickness = value;
                OnPropertyChanged();
            }
        }
    }

    public bool OptimizeDxf
    {
        get => _optimizeDxf;
        set
        {
            if (_optimizeDxf != value)
            {
                _optimizeDxf = value;
                OnPropertyChanged();
            }
        }
    }

    public bool EnableSplineReplacement
    {
        get => _enableSplineReplacement;
        set
        {
            if (_enableSplineReplacement != value)
            {
                _enableSplineReplacement = value;
                OnPropertyChanged();
            }
        }
    }

    public int SelectedSplineReplacementIndex
    {
        get => _selectedSplineReplacementIndex;
        set
        {
            if (_selectedSplineReplacementIndex != value)
            {
                _selectedSplineReplacementIndex = value;
                OnPropertyChanged();
            }
        }
    }

    public int SelectedAcadVersionIndex
    {
        get => _selectedAcadVersionIndex;
        set
        {
            if (_selectedAcadVersionIndex != value)
            {
                _selectedAcadVersionIndex = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsOptimizeDxfEnabled));
                
                // Автоматически отключаем оптимизацию для R12
                if (AcadVersions.Count > value && AcadVersions[value].Value == "R12")
                {
                    OptimizeDxf = false;
                    // Показываем сообщение только если приложение уже инициализировано
                    if (!_isInitializing)
                    {
                        MessageBox.Show("Для версии R12 отключены оптимизация DXF и генерация миниатюр.\n\nДанная версия не поддерживается используемой библиотекой.", 
                            "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
        }
    }

    public ProcessingMethod SelectedProcessingMethod
    {
        get => _selectedProcessingMethod;
        set
        {
            if (_selectedProcessingMethod != value)
            {
                _selectedProcessingMethod = value;
                OnPropertyChanged();
            }
        }
    }

    public bool MergeProfilesIntoPolyline
    {
        get => _mergeProfilesIntoPolyline;
        set
        {
            if (_mergeProfilesIntoPolyline != value)
            {
                _mergeProfilesIntoPolyline = value;
                OnPropertyChanged();
            }
        }
    }

    public bool RebaseGeometry
    {
        get => _rebaseGeometry;
        set
        {
            if (_rebaseGeometry != value)
            {
                _rebaseGeometry = value;
                OnPropertyChanged();
            }
        }
    }

    public bool TrimCenterlines
    {
        get => _trimCenterlines;
        set
        {
            if (_trimCenterlines != value)
            {
                _trimCenterlines = value;
                OnPropertyChanged();
            }
        }
    }

    public ExportFolderType SelectedExportFolder
    {
        get => _selectedExportFolder;
        set
        {
            if (_selectedExportFolder != value)
            {
                _selectedExportFolder = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsSubfolderCheckBoxEnabled)); // Уведомляем о изменении IsEnabled
                UpdateSubfolderOptionsFromProperty();
            }
        }
    }

    public bool EnableSubfolder
    {
        get => _enableSubfolder;
        set
        {
            if (_enableSubfolder != value)
            {
                _enableSubfolder = value;
                OnPropertyChanged();
            }
        }
    }

    // Вычисляемое свойство для IsEnabled чекбокса подпапки
    public bool IsSubfolderCheckBoxEnabled => SelectedExportFolder != ExportFolderType.PartFolder;
    
    public bool EnableFileNameConstructor
    {
        get => _enableFileNameConstructor;
        set
        {
            if (_enableFileNameConstructor != value)
            {
                _enableFileNameConstructor = value;
                OnPropertyChanged();
            }
        }
    }

    public TokenService TokenService => _tokenService;

    // Вычисляемое свойство для доступности оптимизации DXF
    public bool IsOptimizeDxfEnabled => 
        AcadVersions.Count > SelectedAcadVersionIndex && 
        AcadVersions[SelectedAcadVersionIndex].Value != "R12";

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

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
    /// Централизованное управление состоянием UI
    /// </summary>
    private void SetUIState(UIState state)
    {
        ScanButton.IsEnabled = state.ScanEnabled;
        ExportButton.IsEnabled = state.ExportEnabled;
        ClearButton.IsEnabled = state.ClearEnabled;
        ScanButton.Content = state.ScanButtonText;
        ExportButton.Content = state.ExportButtonText;
        ProgressLabelRun.Text = state.ProgressText;
        
        if (state.ProgressValue >= 0)
        {
            if (state.UpdateScanProgress)
                ScanProgressValue = state.ProgressValue;
            if (state.UpdateExportProgress)
                ExportProgressValue = state.ProgressValue;
        }
        
        SetInventorUserInterfaceState(state.InventorUIDisabled);
    }

    /// <summary>
    /// Обновление состояния UI после завершения операции
    /// </summary>
    private void SetUIStateAfterOperation(OperationResult result, OperationType operationType)
    {
        var statusText = result.WasCancelled 
            ? $"Прервано ({GetElapsedTime(result.ElapsedTime)})"
            : operationType == OperationType.Scan
                ? $"Найдено листовых деталей: {result.ProcessedCount} ({GetElapsedTime(result.ElapsedTime)})"
                : $"Завершено ({GetElapsedTime(result.ElapsedTime)})";
        
        var state = UIState.CreateAfterOperationState(_partsData.Count > 0, result.WasCancelled, statusText);
        SetUIState(state);
    }

    /// <summary>
    /// Централизованный показ результатов операции пользователю
    /// </summary>
    private void ShowOperationResult(OperationResult result, OperationType operationType, bool isQuickMode = false)
    {
        if (result.WasCancelled)
        {
            var operationName = operationType == OperationType.Scan ? "сканирования" : "экспорта";
            MessageBox.Show($"Процесс {operationName} был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        switch (operationType)
        {
            case OperationType.Scan:
                // Для сканирования показываем специальные сообщения о конфликтах и ссылках
                if (_conflictingParts.Count > 0)
                {
                    MessageBox.Show($"Обнаружены конфликты обозначений.\nОбщее количество конфликтов: {_conflictingParts.Count}\n\nОбнаружены различные модели или состояния модели с одинаковыми обозначениями. Конфликтующие компоненты исключены из таблицы для предотвращения ошибок.\n\nИспользуйте кнопку \"Анализ обозначений\" на панели инструментов для просмотра деталей конфликтов.",
                        "Конфликт обозначений", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else if (_hasMissingReferences)
                {
                    MessageBox.Show("В сборке обнаружены компоненты с потерянными ссылками. Некоторые данные могли быть пропущены.",
                        "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                break;
                
            case OperationType.Export:
                var exportTitle = isQuickMode ? "Быстрый экспорт DXF завершен" : "Экспорт DXF завершен";
                MessageBox.Show(
                    $"{exportTitle}.\nВсего файлов обработано: {result.ProcessedCount + result.SkippedCount}\nПропущено (без разверток): {result.SkippedCount}\nВсего экспортировано: {result.ProcessedCount}\nВремя выполнения: {GetElapsedTime(result.ElapsedTime)}",
                    "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                break;
        }
    }

    /// <summary>
    /// Централизованная обработка ошибок для операций
    /// </summary>
    private async Task<T> ExecuteWithErrorHandlingAsync<T>(
        Func<Task<T>> operation,
        string operationName) where T : OperationResult, new()
    {
        try
        {
            return await operation();
        }
        catch (OperationCanceledException)
        {
            // Операция была отменена - возвращаем результат с флагом отмены
            var result = new T();
            result.WasCancelled = true;
            return result;
        }
        catch (Exception ex)
        {
            var result = new T();
            result.Errors.Add($"Ошибка {operationName}: {ex.Message}");
            return result;
        }
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
        // Отменяем активные операции и освобождаем ресурсы
        _operationCts?.Cancel();
        _operationCts?.Dispose();
        
        // Сохраняем настройки при закрытии
        SaveSettings();
        
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
        ClearSearchButton.IsEnabled = false;
        SearchTextBox.Focus(); // Устанавливаем фокус на поле поиска
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _actualSearchText = SearchTextBox.Text.Trim().ToLower();

        // Включаем или выключаем кнопку очистки в зависимости от наличия текста
        ClearSearchButton.IsEnabled = !string.IsNullOrEmpty(_actualSearchText);

        // Перезапуск таймера при изменении текста
        _searchDelayTimer.Stop();
        _searchDelayTimer.Start();
    }

    private void AddCustomTextButton_Click(object sender, RoutedEventArgs e)
    {
        var customText = CustomTextBox.Text;
        if (string.IsNullOrEmpty(customText)) return;
        
        // Добавляем текст через TokenizedTextBox
        FileNameTemplateTokenBox.AddCustomText(customText);
        
        // Очищаем поле ввода
        CustomTextBox.Text = string.Empty;
    }

    private void AddSymbolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is string symbol)
        {
            // Добавляем символ через TokenizedTextBox
            FileNameTemplateTokenBox.AddCustomText(symbol);
        }
    }

    private void TemplatePresetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SelectedTemplatePreset != null && !string.IsNullOrEmpty(SelectedTemplatePreset.Template))
        {
            TokenService.FileNameTemplate = SelectedTemplatePreset.Template;
        }
    }

    private void SavePresetInlineButton_Click(object sender, RoutedEventArgs e)
    {
        var currentTemplate = TokenService.FileNameTemplate;
        if (string.IsNullOrEmpty(currentTemplate))
        {
            MessageBox.Show("Шаблон не может быть пустым.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var presetName = PresetNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(presetName))
        {
            MessageBox.Show("Введите имя пресета.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Проверяем, не существует ли уже пресет с таким именем
        var existingPreset = TemplatePresets.FirstOrDefault(p => p.Name == presetName);
        if (existingPreset != null)
        {
            var result = MessageBox.Show(
                $"Пресет с именем '{presetName}' уже существует. Заменить?", 
                "Подтверждение", 
                MessageBoxButton.YesNo, 
                MessageBoxImage.Question);
            
            if (result == MessageBoxResult.Yes)
            {
                existingPreset.Template = currentTemplate;
                PresetNameTextBox.Text = "";
            }
        }
        else
        {
            var newPreset = new TemplatePreset
            {
                Name = presetName,
                Template = currentTemplate
            };
            TemplatePresets.Add(newPreset);
            SelectedTemplatePreset = newPreset;
            PresetNameTextBox.Text = "";
        }
    }

    private void DeletePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTemplatePreset == null)
            return;

        var result = MessageBox.Show(
            $"Удалить пресет '{SelectedTemplatePreset.Name}'?", 
            "Подтверждение удаления", 
            MessageBoxButton.YesNo, 
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            TemplatePresets.Remove(SelectedTemplatePreset);
            SelectedTemplatePreset = null;
        }
    }

    private void CustomTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            AddCustomTextButton_Click(sender, new RoutedEventArgs());
            e.Handled = true;
        }
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
            if (droppedData != null && !PartsDataGrid.Columns.Any(c => c.Header.ToString() == droppedData.ColumnHeader))
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
        if (PartsDataGrid.Columns.Any(c => c.Header.ToString() == iProperty.ColumnHeader))
            return;

        DataGridColumn column;

        // Проверка колонок с шаблонами
        if (ColumnTemplates.TryGetValue(iProperty.ColumnHeader, out var templateName))
        {
            column = new DataGridTemplateColumn
            {
                Header = iProperty.ColumnHeader,
                CellTemplate = FindResource(templateName) as DataTemplate
            };
        }
        else
        {
            // Создаем обычную текстовую колонку через централизованный метод
            column = CreateTextColumn(iProperty.ColumnHeader, iProperty.InventorPropertyName);
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

        // Дозаполняем данные для новой колонки
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
            SetAdornerDeleteMode(true); // Переключаем в режим удаления
        }
        else
        {
            _isColumnDraggedOutside = false;
            SetAdornerDeleteMode(false); // Возвращаем обычный режим
        }
    }

    private void SetAdornerDeleteMode(bool isDeleteMode)
    {
        if (_headerAdorner != null && _headerAdorner.Child is TextBlock textBlock)
        {
            if (isDeleteMode)
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

    /// <summary>
    /// Централизованная валидация активного документа
    /// </summary>
    private DocumentValidationResult ValidateActiveDocument()
    {
        if (!EnsureInventorConnection())
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = "Не удалось подключиться к Inventor"
            };
        }

        var doc = _thisApplication?.ActiveDocument;
        if (doc == null)
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = "Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова."
            };
        }

        var docType = doc.DocumentType switch
        {
            DocumentTypeEnum.kAssemblyDocumentObject => DocumentType.Assembly,
            DocumentTypeEnum.kPartDocumentObject => DocumentType.Part,
            _ => DocumentType.Invalid
        };

        if (docType == DocumentType.Invalid)
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = "Откройте сборку или деталь для работы с приложением."
            };
        }

        return new DocumentValidationResult
        {
            Document = doc,
            DocType = docType,
            IsValid = true,
            DocumentTypeName = docType == DocumentType.Assembly ? "Сборка" : "Деталь"
        };
    }

    /// <summary>
    /// Переподключается к Inventor каждый раз
    /// </summary>
    /// <returns>true если Inventor доступен, false если нет</returns>
    private bool EnsureInventorConnection()
    {
        try
        {
            _thisApplication = (Inventor.Application)MarshalCore.GetActiveObject("Inventor.Application");
            if (_thisApplication != null)
            {
                InitializeProjectAndLibraryData();
                return true;
            }
        }
        catch (COMException)
        {
            MessageBox.Show(
                "Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            _thisApplication = null;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Произошла ошибка при подключении к Inventor: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
            _thisApplication = null;
        }

        return false;
    }

    private void InitializeInventor()
    {
        EnsureInventorConnection();
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
            UpdateUIForCtrlState();
        }
    }

    private void MainWindow_KeyUp(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
        {
            _isCtrlPressed = false;
            UpdateUIForCtrlState();
        }
    }

    private void UpdateUIForCtrlState()
    {
        if (!_isExporting && !_isScanning)
        {
            var currentState = UIState.CreateAfterOperationState(_partsData.Count > 0, false, ProgressLabelRun.Text);
            currentState.ExportEnabled = _isCtrlPressed || _partsData.Count > 0;
            SetUIState(currentState);
        }
    }

    /// <summary>
    /// Обновление прогресса экспорта через централизованную систему UI
    /// </summary>
    private void UpdateExportProgress(double progressValue)
    {
        var progressState = UIState.Exporting;
        progressState.ProgressValue = progressValue;
        progressState.UpdateExportProgress = true;
        progressState.UpdateScanProgress = false;
        SetUIState(progressState);
    }



    private void PrepareExportOptions(out string options)
    {
        var sb = new StringBuilder();

        var selectedAcadVersion = (AcadVersionItem)AcadVersionComboBox.SelectedItem;
        sb.Append($"AcadVersion={selectedAcadVersion?.Value ?? "2000"}");

        if (EnableSplineReplacement)
        {
            var splineTolerance = SplineToleranceTextBox.Text;

            // Получаем текущий разделитель дробной части из системных настроек
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];

            // Заменяем разделитель на актуальный
            splineTolerance = splineTolerance.Replace('.', decimalSeparator).Replace(',', decimalSeparator);

            var selectedSplineType = (SplineReplacementItem)SplineReplacementComboBox.SelectedItem;
            if (selectedSplineType?.Index == 0) // Линии
                sb.Append($"&SimplifySplines=True&SplineTolerance={splineTolerance}");
            else if (selectedSplineType?.Index == 1) // Дуги
                sb.Append($"&SimplifySplines=True&SimplifyAsTangentArcs=True&SplineTolerance={splineTolerance}");
        }
        else
        {
            sb.Append("&SimplifySplines=False");
        }

        if (MergeProfilesIntoPolyline) sb.Append("&MergeProfilesIntoPolyline=True");

        if (RebaseGeometry) sb.Append("&RebaseGeometry=True");

        if (TrimCenterlines) // Проверяем новое поле TrimCenterlinesAtContour
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
    }

    private void IncrementMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out int currentValue))
        {
            currentValue++;
            MultiplierTextBox.Text = currentValue.ToString();
        }
    }

    private void DecrementMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out int currentValue) && currentValue > 1)
        {
            currentValue--;
            MultiplierTextBox.Text = currentValue.ToString();
        }
    }

    /// <summary>
    /// Унифицированный метод сканирования документа
    /// </summary>
    private async Task<OperationResult> ScanDocumentAsync(
        Document document, 
        CancellationToken cancellationToken = default,
        bool updateUI = true,
        IProgress<ScanProgress>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new OperationResult();
        
        try
        {
            var documentType = document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "Сборка" : "Деталь";
            
            if (updateUI)
            {
                // Получаем и отображаем информацию о документе
                UpdateDocumentInfo(documentType, document);
                _partsData.Clear();
                _itemCounter = 1;
            }
            
            _partNumberTracker.Clear();
            ConflictFilesButton.IsEnabled = false;

            var sheetMetalParts = new Dictionary<string, int>();

            if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                var asmDoc = (AssemblyDocument)document;
                
                if (SelectedProcessingMethod == ProcessingMethod.Traverse)
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, progress, cancellationToken), cancellationToken);
                else if (SelectedProcessingMethod == ProcessingMethod.BOM)
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, progress, cancellationToken), cancellationToken);

                // Анализируем конфликты и удаляем конфликтующие детали
                if (updateUI)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        var analysisState = UIState.Scanning;
                        analysisState.ProgressText = "Анализ конфликтов обозначений...";
                        SetUIState(analysisState);
                    });
                }
                
                await AnalyzePartNumberConflictsAsync();
                FilterConflictingParts(sheetMetalParts);

                result.ProcessedCount = sheetMetalParts.Count;

                if (updateUI)
                {
                    // Переключаем прогресс на обработку деталей
                    await Dispatcher.InvokeAsync(() =>
                    {
                        var processingState = UIState.Scanning;
                        processingState.ProgressText = "Обработка деталей...";
                        processingState.ProgressValue = 0;
                        SetUIState(processingState);
                    });

                    var itemCounter = 1;
                    var totalParts = sheetMetalParts.Count;
                    var processedParts = 0;
                    
                    var partProgress = new Progress<PartData>(partData =>
                    {
                        partData.Item = _itemCounter;
                        _partsData.Add(partData);
                        _itemCounter++;
                    });

                    await Task.Run(async () =>
                    {
                        foreach (var part in sheetMetalParts)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var partData = await GetPartDataAsync(part.Key, part.Value, itemCounter++);
                            if (partData != null)
                            {
                                ((IProgress<PartData>)partProgress).Report(partData);
                            }
                            
                            processedParts++;
                            await Dispatcher.InvokeAsync(() =>
                            {
                                var progressState = UIState.Scanning;
                                progressState.ProgressValue = totalParts > 0 ? (double)processedParts / totalParts * 100 : 0;
                                progressState.ProgressText = $"Обработка деталей - {processedParts} из {totalParts}";
                                SetUIState(progressState);
                            });
                        }
                    }, cancellationToken);
                }
            }
            else if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                var partDoc = (PartDocument)document;
                
                if (updateUI)
                {
                    var partProgress = new Progress<PartData>(partData =>
                    {
                        partData.Item = _itemCounter;
                        _partsData.Add(partData);
                        _itemCounter++;
                    });
                    
                    result.ProcessedCount = await ProcessSinglePart(partDoc, partProgress);
                }
                else
                {
                    if (partDoc.SubType == PropertyManager.SheetMetalSubType)
                    {
                        var mgr = new PropertyManager((Document)partDoc);
                        var partNumber = mgr.GetMappedProperty("PartNumber");
                        sheetMetalParts.Add(partNumber, 1);
                        result.ProcessedCount = 1;
                    }
                }
            }

            if (updateUI && int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
                UpdateQuantitiesWithMultiplier(multiplier);

            result.WasCancelled = cancellationToken.IsCancellationRequested;
            
            if (updateUI)
            {
                _lastScannedDocument = cancellationToken.IsCancellationRequested ? null : document;
                _tokenService.UpdatePartsData(_partsData);
            }
        }
        catch (OperationCanceledException)
        {
            result.WasCancelled = true;
        }
        catch (Exception ex)
        {
            result.Errors.Add(ex.Message);
        }
        finally
        {
            stopwatch.Stop();
            result.ElapsedTime = stopwatch.Elapsed;
        }

        return result;
    }

    private async void ScanButton_Click(object sender, RoutedEventArgs e)
    {
        // Если сканирование уже идет, выполняем прерывание
        if (_isScanning)
        {
            _operationCts?.Cancel();
            SetUIState(new UIState { ScanEnabled = false, ScanButtonText = UIState.CANCELLING_TEXT });
            return;
        }

        // Валидация документа
        var validation = ValidateActiveDocument();
        if (!validation.IsValid)
        {
            MessageBox.Show(validation.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Настройка UI для сканирования
        SetUIState(UIState.Scanning);
        _isScanning = true;
        _operationCts = new CancellationTokenSource();
        _hasMissingReferences = false;

        // Прогресс для сканирования структуры сборки
        var scanProgress = new Progress<ScanProgress>(progress =>
        {
            var progressState = UIState.Scanning;
            progressState.ProgressValue = progress.TotalItems > 0 ? (double)progress.ProcessedItems / progress.TotalItems * 100 : 0;
            progressState.ProgressText = $"{progress.CurrentOperation} - {progress.CurrentItem}";
            SetUIState(progressState);
        });

        // Выполнение сканирования
        var result = await ExecuteWithErrorHandlingAsync(
            () => ScanDocumentAsync(validation.Document!, _operationCts.Token, updateUI: true, scanProgress),
            "сканирования");

        // Обработка завершения операции
        _isScanning = false;
        SetUIStateAfterOperation(result, OperationType.Scan);

        // Показ результатов пользователю
        ShowOperationResult(result, OperationType.Scan);
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


    private void UpdateSubfolderOptionsFromProperty()
    {
        // Опция "Рядом с деталью" - чекбокс должен быть отключен
        if (SelectedExportFolder == ExportFolderType.PartFolder)
        {
            EnableSubfolder = false;
        }
        
        // Особая логика для ProjectFolder
        if (SelectedExportFolder == ExportFolderType.ProjectFolder)
        {
            SetProjectFolderInfo(); // Устанавливаем информацию о проекте при выборе
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
            ProjectNameRun.Text = projectName;
            ProjectStackPanelItem.ToolTip = projectWorkspacePath;
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

    private void InitializeAcadVersions()
    {
        AcadVersions = new ObservableCollection<AcadVersionItem>(
            new[] { "2018", "2013", "2010", "2007", "2004", "2000", "R12" }
            .Select(v => new AcadVersionItem { DisplayName = v, Value = v })
        );
    }

    private void InitializeSplineReplacementTypes()
    {
        SplineReplacementTypes = new ObservableCollection<SplineReplacementItem>
        {
            new() { DisplayName = "Линии", Index = 0 },
            new() { DisplayName = "Дуги", Index = 1 }
        };
    }

    private void InitializeAvailableTokens()
    {
        AvailableTokens = new ObservableCollection<string>
        {
            "PartNumber",
            "Quantity",
            "Material",
            "Thickness",
            "Description",
            "Author",
            "Revision",
            "Project",
            "Mass",
            "FlatPatternWidth",
            "FlatPatternLength",
            "FlatPatternArea"
        };
    }

    private void AddPartToConflictTracker(string partNumber, string fileName, string modelState)
    {
        var conflictInfo = new PartConflictInfo
        {
            PartNumber = partNumber,
            FileName = fileName,
            ModelState = modelState
        };
        
        _partNumberTracker.AddOrUpdate(partNumber,
            new List<PartConflictInfo> { conflictInfo },
            (key, existingList) =>
            {
                // Проверяем, есть ли уже такой же файл + состояние модели
                if (!existingList.Any(p => p.UniqueId == conflictInfo.UniqueId))
                {
                    existingList.Add(conflictInfo);
                }
                return existingList;
            });
    }

    private async Task<int> ProcessSinglePart(PartDocument partDoc, IProgress<PartData> partProgress)
    {
        if (partDoc.SubType == PropertyManager.SheetMetalSubType)
        {
            var processingState = UIState.Scanning;
            processingState.ProgressText = "Обработка детали...";
            SetUIState(processingState);
            
            var partData = await GetPartDataAsync(partDoc, 1, 1);
            if (partData != null)
            {
                ((IProgress<PartData>)partProgress).Report(partData);
                
                var completedState = UIState.Scanning;
                completedState.ProgressValue = 100;
                SetUIState(completedState);
                return 1; // Успешно обработана 1 деталь
            }
        }
        return 0; // Деталь не обработана (не листовой металл или ошибка)
    }

    private void FilterConflictingParts(Dictionary<string, int> sheetMetalParts)
    {
        // Получаем список конфликтующих обозначений
        var conflictingPartNumbers = _partNumberTracker
            .Where(p => p.Value.Count > 1)
            .Select(p => p.Key)
            .ToHashSet();

        // Удаляем конфликтующие детали из sheetMetalParts
        foreach (var conflictingPartNumber in conflictingPartNumbers)
        {
            sheetMetalParts.Remove(conflictingPartNumber);
        }
    }

    private async Task AnalyzePartNumberConflictsAsync()
    {
        _conflictingParts.Clear(); // Очищаем список конфликтов перед началом проверки

        // Оставляем только те записи, у которых больше одного файла (т.е. есть конфликт)
        var conflictingPartNumbers = _partNumberTracker.Where(p => p.Value.Count > 1).ToDictionary(p => p.Key, p => p.Value);

        if (conflictingPartNumbers.Any())
        {
            _conflictingParts.AddRange(conflictingPartNumbers.SelectMany(entry => entry.Value.Select(v => new PartData { PartNumber = v.PartNumber })));

            // Используем Dispatcher для обновления UI без показа MessageBox
            await Dispatcher.InvokeAsync(() =>
            {
                ConflictFilesButton.IsEnabled = true; // Включаем кнопку для просмотра подробностей о конфликтах
                _conflictFileDetails = conflictingPartNumbers; // Сохраняем подробную информацию о конфликтах для вывода при нажатии кнопки
            });
        }
    }


    private System.Windows.Point _startPoint;
    private bool _isDragging = false;

    private void AvailableTokensListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _startPoint = e.GetPosition(null);
        _isDragging = false;
    }

    private void AvailableTokensListBox_MouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
        {
            var currentPosition = e.GetPosition(null);
            if ((Math.Abs(currentPosition.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance) ||
                (Math.Abs(currentPosition.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                var listBox = sender as System.Windows.Controls.ListBox;
                if (listBox?.SelectedItem is string selectedToken)
                {
                    _isDragging = true;
                    System.Windows.DragDrop.DoDragDrop(listBox, $"{{{selectedToken}}}", System.Windows.DragDropEffects.Copy);
                    _isDragging = false;
                }
            }
        }
    }

    private void AvailableTokensListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var listBox = sender as System.Windows.Controls.ListBox;
        if (listBox?.SelectedItem is string selectedToken)
        {
            FileNameTemplateTokenBox.AddToken($"{{{selectedToken}}}");
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
    public BitmapImage? Preview { get; set; } // apprenticeDoc.Thumbnail
    public bool HasFlatPattern { get; set; } = false; // smCompDef.HasFlatPattern


    // === КАТЕГОРИЯ 2/4: ОБЯЗАТЕЛЬНЫЕ СВОЙСТВА БЕЗ УВЕДОМЛЕНИЙ ===
    // Свойства, которые устанавливаются только при создании объекта
    public string Material { get; set; } = string.Empty;
    public double Thickness { get; set; } = 0.0;

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
    
    private bool isOverridden;
    public bool IsOverridden 
    { 
        get => isOverridden; 
        set 
        { 
            isOverridden = value; 
            OnPropertyChanged(); 
        } 
    }
    
    private bool isMultiplied;
    public bool IsMultiplied 
    { 
        get => isMultiplied; 
        set 
        { 
            isMultiplied = value; 
            OnPropertyChanged(); 
        } 
    }

    // === КАТЕГОРИЯ 7: СВОЙСТВА СОСТОЯНИЯ ОБРАБОТКИ ===
    // Свойства, устанавливаемые во время экспорта и обработки
    private BitmapImage? dxfPreview;
    public BitmapImage? DxfPreview 
    { 
        get => dxfPreview; 
        set 
        { 
            if (dxfPreview != value)
            {
                dxfPreview = value; 
                OnPropertyChanged(); 
            }
        } 
    }
    
    private ProcessingStatus processingStatus = ProcessingStatus.NotProcessed;
    public ProcessingStatus ProcessingStatus 
    { 
        get => processingStatus; 
        set 
        { 
            if (processingStatus != value)
            {
                processingStatus = value; 
                OnPropertyChanged(); 
            }
        } 
    }

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

    // КАТЕГОРИЯ 2/4: Обязательные свойства с уведомлениями (изменяются программно)
    private string partNumber = string.Empty;
    private string description = string.Empty;
    private bool partNumberIsExpression = false;

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

    public bool PartNumberIsExpression
    {
        get => partNumberIsExpression;
        set
        {
            if (partNumberIsExpression != value)
            {
                partNumberIsExpression = value;
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

    public void OnPropertyChanged([CallerMemberName] string? name = null)
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
        customProperties.Remove(propertyName);
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
    public string ColumnHeader { get; set; } = string.Empty; // Заголовок колонки в DataGrid (например, "Обозначение")
    public string ListDisplayName { get; set; } = string.Empty; // Отображение в списке выбора свойств (например, "Обозначение")
    public string InventorPropertyName { get; set; } = string.Empty; // Соответствующее имя свойства iProperty в Inventor (например, "PartNumber")
    public string Category { get; set; } = string.Empty; // Категория свойства для группировки
}

// Структура для отслеживания прогресса сканирования
public class ScanProgress
{
    public int ProcessedItems { get; set; }
    public int TotalItems { get; set; }
    public string CurrentOperation { get; set; } = string.Empty;
    public string CurrentItem { get; set; } = string.Empty;
}


// Конвертер для работы с enum в RadioButton
public class EnumToBooleanConverter : IValueConverter
{
    public static readonly EnumToBooleanConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null)
            return false;

        // Универсальный подход для любых enum
        var valueType = value.GetType();
        if (valueType.IsEnum && Enum.TryParse(valueType, parameter.ToString(), out var parameterValue))
        {
            return value.Equals(parameterValue);
        }

        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue && boolValue && parameter != null && targetType.IsEnum)
        {
            if (Enum.TryParse(targetType, parameter.ToString(), out var result))
            {
                return result;
            }
        }

        return Binding.DoNothing;
    }
}

public class ObjectToBooleanConverter : IValueConverter
{
    public static readonly ObjectToBooleanConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value != null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

public class TemplatePreset
{
    public string Name { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
    
    public override string ToString() => Name;
}

// Структура для детальной информации о конфликтах обозначений
public class PartConflictInfo
{
    public string PartNumber { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string ModelState { get; set; } = string.Empty;
    
    // Уникальный идентификатор для сравнения
    public string UniqueId => $"{FileName}|{ModelState}";
    
    public override bool Equals(object? obj)
    {
        return obj is PartConflictInfo other && UniqueId == other.UniqueId;
    }
    
    public override int GetHashCode()
    {
        return UniqueId.GetHashCode();
    }
}

// Enum для статуса обработки экспорта
public enum ProcessingStatus
{
    NotProcessed,   // Не обработан (прозрачный)
    Success,        // Успешно экспортирован (зеленый)
    Error          // Ошибка экспорта (красный)
}

// Enum для типа документа
public enum DocumentType
{
    Assembly,
    Part,
    Invalid
}

// Enum для типа операции
public enum OperationType
{
    Scan,
    Export
}

// Класс для управления состоянием UI
public class UIState
{
    private const string SCAN_BUTTON_TEXT = "Сканировать";
    private const string CANCEL_BUTTON_TEXT = "Прервать";
    private const string EXPORT_BUTTON_TEXT = "Экспорт";
    public const string CANCELLING_TEXT = "Прерывание...";
    
    public bool ScanEnabled { get; set; }
    public bool ExportEnabled { get; set; }
    public bool ClearEnabled { get; set; }
    public string ScanButtonText { get; set; } = "";
    public string ExportButtonText { get; set; } = "";
    public string ProgressText { get; set; } = "";
    public double ProgressValue { get; set; }
    public bool InventorUIDisabled { get; set; }
    public bool UpdateScanProgress { get; set; } = true;
    public bool UpdateExportProgress { get; set; } = true;
    
    public static UIState Initial => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "Документ не выбран",
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = true
    };
    
    public static UIState Scanning => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = CANCEL_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "Подготовка к сканированию...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = true,
        UpdateExportProgress = false
    };
    
    public static UIState Exporting => new()
    {
        ScanEnabled = false,
        ExportEnabled = true,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = CANCEL_BUTTON_TEXT,
        ProgressText = "Экспорт данных...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = false,
        UpdateExportProgress = true
    };
    
    public static UIState CreateClearedState() => new()
    {
        ScanEnabled = true,
        ExportEnabled = false,
        ClearEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = "",
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = false
    };
    
    public static UIState CreateAfterOperationState(bool hasData, bool wasCancelled, string statusText) => new()
    {
        ScanEnabled = true,
        ExportEnabled = hasData && !wasCancelled,
        ClearEnabled = hasData,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ProgressText = statusText,
        ProgressValue = 0,
        InventorUIDisabled = false,
        UpdateScanProgress = true,
        UpdateExportProgress = true
    };
}

// Контекст для экспорта
public class ExportContext
{
    public string TargetDirectory { get; set; } = "";
    public int Multiplier { get; set; } = 1;
    public Dictionary<string, int> SheetMetalParts { get; set; } = new();
    public bool GenerateThumbnails { get; set; } = true;
    public bool IsValid { get; set; } = true;
    public string ErrorMessage { get; set; } = "";
}

// Результат операции
public class OperationResult
{
    public int ProcessedCount { get; set; }
    public int SkippedCount { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public bool WasCancelled { get; set; }
    public List<string> Errors { get; set; } = new();
    public bool IsSuccess => !WasCancelled && Errors.Count == 0;
}

// Результат валидации документа
public class DocumentValidationResult
{
    public Document? Document { get; set; }
    public DocumentType DocType { get; set; }
    public bool IsValid { get; set; }
    public string ErrorMessage { get; set; } = "";
    public string DocumentTypeName { get; set; } = "";
}