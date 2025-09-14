using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
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
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;

namespace FlatPatternExporter;


public class AcadVersionItem
{
    public string DisplayName { get; set; } = string.Empty;
    public AcadVersionType Value { get; set; }
    public override string ToString() => DisplayName;
}

public class SplineReplacementItem
{
    public string DisplayName { get; set; } = string.Empty;
    public SplineReplacementType Value { get; set; }
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
    private readonly ObservableCollection<PartData> _partsData = [];
    private readonly CollectionViewSource _partsDataView;
    private List<PartData> _conflictingParts = new();
    private Dictionary<string, List<PartConflictInfo>> _conflictFileDetails = new();
    private readonly ConcurrentDictionary<string, List<PartConflictInfo>> _partNumberTracker = new();
    
    // Кеш документов для изоляции среды обработки
    private readonly Dictionary<string, PartDocument> _documentCache = new();
    private readonly Dictionary<string, string> _partNumberToFullFileName = new();
    
    // Состояние процессов
    private bool _isScanning;
    private bool _isExporting;
    private CancellationTokenSource? _operationCts;
    private bool _hasMissingReferences = false;
    private int _itemCounter = 1;
    private bool _isInitializing = true;
    private bool _isUpdatingState;
    
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
    private SplineReplacementType _selectedSplineReplacement = SplineReplacementType.Lines;
    private AcadVersionType _selectedAcadVersion = AcadVersionType.V2000; // 2000 по умолчанию
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
    
    // Метод обработки
    private ProcessingMethod _selectedProcessingMethod = ProcessingMethod.BOM;
    
    // Модель состояния
    private bool _isPrimaryModelState = true;
    
    // Словарь колонок с шаблонами (все типы)
    private static readonly Dictionary<string, string> ColumnTemplates = InitializeColumnTemplates();

    private static Dictionary<string, string> InitializeColumnTemplates()
    {
        var templates = new Dictionary<string, string>();

        // Автоматически добавляем шаблоны для всех свойств из централизованного реестра
        foreach (var definition in PropertyMetadataRegistry.Properties.Values)
        {
            var columnName = definition.ColumnHeader.Length > 0 ? definition.ColumnHeader : definition.DisplayName;
            
            // Используем уникальный шаблон, если он задан
            if (!string.IsNullOrEmpty(definition.ColumnTemplate))
            {
                templates[columnName] = definition.ColumnTemplate;
            }
            // Для редактируемых свойств без уникального шаблона создаем шаблон с FX индикатором
            else if (definition.IsEditable)
            {
                templates[columnName] = $"{definition.InternalName}WithExpressionTemplate";
            }
        }

        return templates;
    }



    public FlatPatternExporterMainWindow()
    {
        InitializeComponent();
        InitializeInventor();
        SetProjectFolderInfo(); // Инициализация папки проекта при запуске
        PartsDataGrid.ItemsSource = _partsData;
        PartsDataGrid.PreviewMouseMove += PartsDataGrid_PreviewMouseMove;
        PartsDataGrid.PreviewMouseLeftButtonUp += PartsDataGrid_PreviewMouseLeftButtonUp;
        PartsDataGrid.ColumnReordering += PartsDataGrid_ColumnReordering;
        PartsDataGrid.ColumnReordered += PartsDataGrid_ColumnReordered;
        
        // Обновляем видимость оверлея при изменении коллекции колонок
        ((System.Collections.Specialized.INotifyCollectionChanged)PartsDataGrid.Columns).CollectionChanged += (s, e) => UpdateNoColumnsOverlayVisibility();
        

        // Настройка TokenService для работы с визуальным контейнером
        TokenService?.SetTokenContainer(TokenContainer);
        
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
        PresetIProperties = PropertyMetadataRegistry.GetPresetProperties();

        // Устанавливаем DataContext для текущего окна
        DataContext = this;

        // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
        KeyDown += MainWindow_KeyDown;
        KeyUp += MainWindow_KeyUp;

        PresetManager.PropertyChanged += PresetManager_PropertyChanged;
        PresetManager.TemplatePresets.CollectionChanged += TemplatePresets_CollectionChanged;
        TokenService!.PropertyChanged += TokenService_PropertyChanged;

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
    public ObservableCollection<PropertyMetadataRegistry.PropertyDefinition> AvailableTokens { get; set; } = new();
    public ObservableCollection<PropertyMetadataRegistry.PropertyDefinition> UserDefinedTokens { get; set; } = new();
    public TemplatePresetManager PresetManager { get; } = new();    

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

    public SplineReplacementType SelectedSplineReplacement
    {
        get => _selectedSplineReplacement;
        set
        {
            if (_selectedSplineReplacement != value)
            {
                _selectedSplineReplacement = value;
                OnPropertyChanged();
            }
        }
    }

    public AcadVersionType SelectedAcadVersion
    {
        get => _selectedAcadVersion;
        set
        {
            if (_selectedAcadVersion != value)
            {
                _selectedAcadVersion = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsOptimizeDxfEnabled));

                if (!AcadVersionMapping.SupportsOptimization(_selectedAcadVersion))
                {
                    OptimizeDxf = false;
                    if (!_isInitializing)
                    {
                        var ver = AcadVersionMapping.GetDisplayName(_selectedAcadVersion);
                        MessageBox.Show(
                            $"Для версии {ver} отключены оптимизация DXF и генерация миниатюр.\n\nДанная версия не поддерживается используемой библиотекой.",
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
    public bool IsOptimizeDxfEnabled => AcadVersionMapping.SupportsOptimization(SelectedAcadVersion);

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

    private bool _isRenameButtonEnabled;
    public bool IsRenameButtonEnabled
    {
        get => _isRenameButtonEnabled;
        set
        {
            if (_isRenameButtonEnabled != value)
            {
                _isRenameButtonEnabled = value;
                OnPropertyChanged();
            }
        }
    }

    private bool _isPresetSelected;
    public bool IsPresetSelected
    {
        get => _isPresetSelected;
        set
        {
            if (_isPresetSelected != value)
            {
                _isPresetSelected = value;
                OnPropertyChanged();
            }
        }
    }

    private bool _isSaveButtonEnabled;
    public bool IsSaveButtonEnabled
    {
        get => _isSaveButtonEnabled;
        set
        {
            if (_isSaveButtonEnabled != value)
            {
                _isSaveButtonEnabled = value;
                OnPropertyChanged();
            }
        }
    }

    private bool _isCreateButtonEnabled;
    public bool IsCreateButtonEnabled
    {
        get => _isCreateButtonEnabled;
        set
        {
            if (_isCreateButtonEnabled != value)
            {
                _isCreateButtonEnabled = value;
                OnPropertyChanged();
            }
        }
    }

    private void TemplatePresets_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        UpdateEmptyState();
    }

    private void PresetManager_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TemplatePresetManager.SelectedTemplatePreset))
        {
            UpdatePresetNameTextBox();
            UpdateButtonStates();
        }
    }

    private void TokenService_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TokenService.FileNameTemplate) ||
            e.PropertyName == nameof(TokenService.IsFileNameTemplateValid))
        {
            UpdateButtonStates();
        }
    }

    private void UpdateEmptyState()
    {
        var isEmpty = PresetManager?.TemplatePresets.Count == 0;
        if (EmptyStatePanel != null && TemplatePresetsListBox != null)
        {
            EmptyStatePanel.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
            TemplatePresetsListBox.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private void UpdatePresetNameTextBox()
    {
        var selected = PresetManager?.SelectedTemplatePreset;
        _isUpdatingState = true;
        try
        {
            if (PresetNameTextBox != null)
            {
                PresetNameTextBox.Text = selected?.Name ?? string.Empty;
            }
            if (selected != null && TokenService != null && TokenService.FileNameTemplate != selected.Template)
            {
                TokenService.FileNameTemplate = selected.Template;
            }
        }
        finally
        {
            _isUpdatingState = false;
        }
    }

    private void PresetNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (!_isUpdatingState)
        {
            UpdateButtonStates();
        }
    }

    private void UpdateButtonStates()
    {
        if (_isUpdatingState) return;

        var selected = PresetManager?.SelectedTemplatePreset;
        var trimmedName = GetTrimmedPresetName();

        IsPresetSelected = selected != null;

        if (selected != null)
        {
            bool nameChanged = !string.Equals(trimmedName, selected.Name, StringComparison.Ordinal);
            bool duplicateExists = PresetManager!.TemplatePresets
                .Any(p => !ReferenceEquals(p, selected) && string.Equals(p.Name, trimmedName, StringComparison.OrdinalIgnoreCase));

            IsRenameButtonEnabled = !string.IsNullOrWhiteSpace(trimmedName) && nameChanged && !duplicateExists;
        }
        else
        {
            IsRenameButtonEnabled = false;
        }

        IsSaveButtonEnabled = selected != null &&
                              TokenService != null &&
                              selected.Template != TokenService.FileNameTemplate;

        IsCreateButtonEnabled = PresetManager != null &&
                                TokenService != null &&
                                !string.IsNullOrWhiteSpace(trimmedName) &&
                                !PresetManager.PresetNameExists(trimmedName) &&
                                !string.IsNullOrWhiteSpace(TokenService.FileNameTemplate) &&
                                TokenService.IsFileNameTemplateValid;
    }

    private void DeletePresetButton_Click(object sender, RoutedEventArgs e)
    {
        PresetManager?.DeleteSelectedPreset();
    }

    private void CreatePresetButton_Click(object sender, RoutedEventArgs e)
    {
        PresetManager?.CreatePreset(GetTrimmedPresetName(), TokenService!.FileNameTemplate, out _);
    }

    private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset() || TokenService == null) return;
        PresetManager!.UpdateSelectedTemplate(TokenService.FileNameTemplate);
        UpdateButtonStates();
    }

    private void RenamePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;
        PresetManager!.RenameSelected(GetTrimmedPresetName(), out _);
    }

    private void DuplicatePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;
        PresetManager!.DuplicateSelected();
    }

    private bool HasSelectedPreset() => PresetManager?.SelectedTemplatePreset != null;
    private string GetTrimmedPresetName() => (PresetNameTextBox?.Text ?? string.Empty).Trim();


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
                    MessageBox.Show($"Обнаружены конфликты обозначений.\nОбщее количество конфликтов: {_conflictingParts.Count}\n\nОбнаружены различные модели или состояния модели с одинаковыми обозначениями. Конфликтующие компоненты исключены из таблицы для предотвращения ошибок.\n\nИспользуйте кнопку \"Конфликты\" на панели управления для просмотра деталей конфликтов.",
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

    /// <summary>
    /// Централизованная валидация активного документа с показом ошибки
    /// </summary>
    private DocumentValidationResult? ValidateDocumentOrShowError()
    {
        var validation = ValidateActiveDocument();
        if (!validation.IsValid)
        {
            MessageBox.Show(validation.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }
        return validation;
    }

    /// <summary>
    /// Централизованная подготовка контекста экспорта с показом ошибки
    /// </summary>
    private async Task<ExportContext?> PrepareExportContextOrShowError(Document document, bool requireScan = true, bool showProgress = false)
    {
        var context = await PrepareExportContextAsync(document, requireScan, showProgress);
        if (!context.IsValid)
        {
            MessageBox.Show(context.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            if (context.ErrorMessage.Contains("сканирование"))
                MultiplierTextBox.Text = "1";
            return null;
        }
        return context;
    }

    /// <summary>
    /// Централизованная инициализация операции
    /// </summary>
    private void InitializeOperation(UIState operationState, ref bool operationFlag)
    {
        SetUIState(operationState);
        operationFlag = true;
        _operationCts = new CancellationTokenSource();
    }

    /// <summary>
    /// Централизованное завершение операции
    /// </summary>
    private void CompleteOperation(OperationResult result, OperationType operationType, ref bool operationFlag, bool isQuickMode = false)
    {
        operationFlag = false;
        SetUIStateAfterOperation(result, operationType);
        ShowOperationResult(result, operationType, isQuickMode);
    }

    /// <summary>
    /// Централизованное создание результата экспортной операции
    /// </summary>
    private OperationResult CreateExportOperationResult(int processedCount, int skippedCount, TimeSpan elapsedTime)
    {
        return new OperationResult
        {
            ProcessedCount = processedCount,
            SkippedCount = skippedCount,
            ElapsedTime = elapsedTime,
            WasCancelled = _operationCts!.Token.IsCancellationRequested
        };
    }

    /// <summary>
    /// Централизованная обработка прерывания экспорта
    /// </summary>
    private bool HandleExportCancellation()
    {
        if (_isExporting)
        {
            _operationCts?.Cancel();
            SetUIState(UIState.CancellingExport);
            return true;
        }
        return false;
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
        var selectIPropertyWindow = new SelectIPropertyWindow(
            PresetIProperties,
            columnHeader => PartsDataGrid.Columns.Any(c => c.Header.ToString() == columnHeader),
            AddIPropertyColumn,
            AddUserDefinedIPropertyColumn,
            RemoveDataGridColumn);
        selectIPropertyWindow.ShowDialog();
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
        
        // Добавляем текст через TokenService
        TokenService?.AddCustomText(customText);
        
        // Очищаем поле ввода
        CustomTextBox.Text = string.Empty;
    }

    private void AddSymbolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is string symbol)
        {
            // Добавляем символ через TokenService
            TokenService?.AddCustomText(symbol);
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
        if (e.Item is not PartData partData)
        {
            e.Accepted = false;
            return;
        }

        if (string.IsNullOrEmpty(_actualSearchText))
        {
            e.Accepted = true;
            return;
        }

        // Проходим по всем видимым колонкам DataGrid
        foreach (var column in PartsDataGrid.Columns)
        {
            var columnHeader = column.Header.ToString();
            if (string.IsNullOrEmpty(columnHeader)) continue;

            // Ищем по стандартным свойствам
            var metadata = PropertyMetadataRegistry.Properties.Values
                .FirstOrDefault(p => p.ColumnHeader == columnHeader);
            
            if (metadata != null && !metadata.IsSearchable) continue;

            string? searchValue = null;

            // Получаем значение из стандартного свойства
            if (metadata != null)
            {
                var propInfo = typeof(PartData).GetProperty(metadata.InternalName);
                searchValue = propInfo?.GetValue(partData)?.ToString();
            }
            // Получаем значение из пользовательского свойства
            // Находим User Defined Property по ColumnHeader
            var userProp = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeader);
            if (userProp != null && partData.UserDefinedProperties.ContainsKey(userProp.InventorPropertyName!))
            {
                searchValue = partData.UserDefinedProperties[userProp.InventorPropertyName!];
            }

            // Проверяем совпадение
            if (!string.IsNullOrEmpty(searchValue) && 
                searchValue.ToLower().Contains(_actualSearchText))
            {
                e.Accepted = true;
                return;
            }
        }

        e.Accepted = false;
    }

    public void AddIPropertyColumn(PresetIProperty iProperty)
    {
        // Проверяем, существует ли уже колонка с таким заголовком
        if (PartsDataGrid.Columns.Any(c => c.Header.ToString() == iProperty.ColumnHeader))
            return;

        // Получаем метаданные свойства для проверки возможности сортировки
        var isSortable = true;
        if (PropertyMetadataRegistry.Properties.TryGetValue(iProperty.InventorPropertyName, out var propertyDef))
        {
            isSortable = propertyDef.IsSortable;
        }

        DataGridColumn column;

        // Проверка колонок с шаблонами
        if (ColumnTemplates.TryGetValue(iProperty.ColumnHeader, out var templateName))
        {
            DataTemplate template;
            
            // Если это редактируемое свойство с fx индикатором - создаем динамически
            if (templateName.EndsWith("WithExpressionTemplate"))
            {
                template = FindResource("EditableWithFxTemplate") as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException("EditableWithFxTemplate не найден", "EditableWithFxTemplate");
            }
            else
            {
                // Для остальных шаблонов ищем в ресурсах
                template = FindResource(templateName) as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException($"Шаблон {templateName} не найден", templateName);
            }
            
            var templateColumn = new DataGridTemplateColumn
            {
                Header = iProperty.ColumnHeader,
                CellTemplate = template,
                SortMemberPath = isSortable ? iProperty.InventorPropertyName : null,
                IsReadOnly = !templateName.StartsWith("Editable"),
                ClipboardContentBinding = new Binding(iProperty.InventorPropertyName)
            };

            // Передаём путь свойства через Tag ячейки для универсального шаблона
            if (templateName.EndsWith("WithExpressionTemplate"))
            {
                templateColumn.CellStyle = CreateCellTagStyle(iProperty.InventorPropertyName);
            }

            column = templateColumn;
        }
        else
        {
            // Создаем обычную текстовую колонку через централизованный метод
            column = CreateTextColumn(iProperty.ColumnHeader, iProperty.InventorPropertyName, isSortable);
        }

        // Добавляем колонку в конец
        PartsDataGrid.Columns.Add(column);

        // Обновляем состояние свойств в окне выбора
        var selectIPropertyWindow =
            System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
        selectIPropertyWindow?.UpdatePropertyStates();

        // Дозаполняем данные для новой колонки
        FillPropertyData(iProperty.InventorPropertyName);
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

            // Используем централизованный метод с removeDataOnly=true для Adorner
            RemoveDataGridColumn(columnName, removeDataOnly: true);
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
                InitializeProjectData();
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
    private void InitializeProjectData()
    {
        if (_thisApplication != null)
        {
            try
            {
                SetProjectFolderInfo();
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

        sb.Append($"AcadVersion={AcadVersionMapping.GetTranslatorCode(SelectedAcadVersion)}");

        if (EnableSplineReplacement)
        {
            var splineTolerance = SplineToleranceTextBox.Text;

            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
            splineTolerance = splineTolerance.Replace('.', decimalSeparator).Replace(',', decimalSeparator);

            if (SelectedSplineReplacement == SplineReplacementType.Lines)
                sb.Append($"&SimplifySplines=True&SplineTolerance={splineTolerance}");
            else if (SelectedSplineReplacement == SplineReplacementType.Arcs)
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
            partData.IsOverridden = false;
            partData.SetQuantityInternal(partData.OriginalQuantity * multiplier);
            partData.IsMultiplied = multiplier > 1;
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
            
            ClearConflictData();

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
                
                // Проверяем, не является ли деталь библиотечным компонентом
                if (!IncludeLibraryComponents && IsLibraryComponent(partDoc.FullFileName))
                {
                    result.ProcessedCount = 0;
                    return result;
                }
                
                var mgr = new PropertyManager((Document)partDoc);
                var partNumber = mgr.GetMappedProperty("PartNumber");
                if (!string.IsNullOrEmpty(partNumber))
                {
                    // Добавляем документ в кеш
                    AddDocumentToCache(partDoc, partNumber);
                }
                
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
            SetUIState(UIState.CancellingScan);
            return;
        }

        // Валидация документа
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Настройка UI для сканирования
        InitializeOperation(UIState.Scanning, ref _isScanning);
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
            () => ScanDocumentAsync(validation.Document!, _operationCts!.Token, updateUI: true, scanProgress),
            "сканирования");

        // Завершение операции
        CompleteOperation(result, OperationType.Scan, ref _isScanning);
    }
    private void ConflictFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_conflictFileDetails == null || _conflictFileDetails.Count == 0)
        {
            return;
        }

        // Создаем и показываем окно с деталями конфликтов
        var conflictWindow = new ConflictDetailsWindow(_conflictFileDetails, OpenInventorDocument);
        conflictWindow.Owner = this;
        conflictWindow.ShowDialog();
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

    private void InitializeAcadVersions()
    {
        AcadVersions = new ObservableCollection<AcadVersionItem>(
            Enum.GetValues<AcadVersionType>()
                .Select(version => new AcadVersionItem
                {
                    DisplayName = AcadVersionMapping.GetDisplayName(version),
                    Value = version
                })
        );
    }

    private void InitializeSplineReplacementTypes()
    {
        SplineReplacementTypes = new ObservableCollection<SplineReplacementItem>(
            Enum.GetValues<SplineReplacementType>()
                .Select(type => new SplineReplacementItem
                {
                    DisplayName = SplineReplacementMapping.GetDisplayName(type),
                    Value = type
                })
        );
    }

    private void InitializeAvailableTokens()
    {
        AvailableTokens = new ObservableCollection<PropertyMetadataRegistry.PropertyDefinition>();
        UserDefinedTokens = new ObservableCollection<PropertyMetadataRegistry.PropertyDefinition>();
        
        RefreshAvailableTokens();
        
        // Подписываемся на изменения в коллекции пользовательских свойств
        PropertyMetadataRegistry.UserDefinedProperties.CollectionChanged += (s, e) =>
        {
            RefreshAvailableTokens();
        };
    }
    
    private void RefreshAvailableTokens()
    {
        AvailableTokens.Clear();
        UserDefinedTokens.Clear();
        
        // Получаем все токенизируемые свойства из централизованного реестра
        var tokenizableProperties = PropertyMetadataRegistry.GetTokenizableProperties();
        
        // Разделяем на стандартные и пользовательские свойства
        var standardProperties = tokenizableProperties.Where(p => p.Type != PropertyMetadataRegistry.PropertyType.UserDefined).OrderBy(p => p.DisplayName);
        var userDefinedProperties = tokenizableProperties.Where(p => p.Type == PropertyMetadataRegistry.PropertyType.UserDefined).OrderBy(p => p.DisplayName);
        
        // Добавляем стандартные свойства
        foreach (var property in standardProperties)
        {
            AvailableTokens.Add(property);
        }
        
        // Добавляем пользовательские свойства
        foreach (var property in userDefinedProperties)
        {
            UserDefinedTokens.Add(property);
        }
    }

    /// <summary>
    /// Очищает все коллекции связанные с конфликтами для изоляции операций
    /// </summary>
    private void ClearConflictData()
    {
        _partNumberTracker.Clear();
        _conflictingParts.Clear();
        _conflictFileDetails.Clear();
        ConflictFilesButton.IsEnabled = false; // Отключаем кнопку при очистке данных
    }

    private void AddDocumentToCache(PartDocument partDoc, string partNumber)
    {
        if (!string.IsNullOrEmpty(partNumber) && !_documentCache.ContainsKey(partNumber))
        {
            _documentCache[partNumber] = partDoc;
            _partNumberToFullFileName[partNumber] = partDoc.FullFileName;
        }
    }

    private PartDocument? GetCachedPartDocument(string partNumber)
    {
        return _documentCache.TryGetValue(partNumber, out var partDoc) ? partDoc : null;
    }

    private void ClearDocumentCache()
    {
        _documentCache.Clear();
        _partNumberToFullFileName.Clear();
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

            // Используем Dispatcher для обновления UI
            await Dispatcher.InvokeAsync(() =>
            {
                ConflictFilesButton.IsEnabled = true; // Включаем кнопку для просмотра подробностей о конфликтах
                _conflictFileDetails = conflictingPartNumbers; // Сохраняем подробную информацию о конфликтах для вывода при нажатии кнопки
            });
        }
    }

    private void TokenListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var listBox = sender as System.Windows.Controls.ListBox;
        if (listBox?.SelectedItem is PropertyMetadataRegistry.PropertyDefinition selectedProperty)
        {
            TokenService?.AddToken($"{{{selectedProperty.TokenName}}}");
        }
    }
}

public class PartData : INotifyPropertyChanged
{
    private int item;

    public string FileName { get; set; } = "";
    public string FullFileName { get; set; } = "";
    public string ModelState { get; set; } = "";
    public BitmapImage? Preview { get; set; }
    public bool HasFlatPattern { get; set; }
    public string Material { get; set; } = "";
    public string Thickness { get; set; } = "";
    public string PartNumber { get; set; } = "";
    public string Description { get; set; } = "";
    public string Author { get; set; } = "";
    public string Revision { get; set; } = "";
    public string Project { get; set; } = "";
    public string StockNumber { get; set; } = "";
    public string Title { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Keywords { get; set; } = "";
    public string Comments { get; set; } = "";
    public string Category { get; set; } = "";
    public string Manager { get; set; } = "";
    public string Company { get; set; } = "";
    public string CreationTime { get; set; } = "";
    public string CostCenter { get; set; } = "";
    public string CheckedBy { get; set; } = "";
    public string EngApprovedBy { get; set; } = "";
    public string UserStatus { get; set; } = "";
    public string CatalogWebLink { get; set; } = "";
    public string Vendor { get; set; } = "";
    public string MfgApprovedBy { get; set; } = "";
    public string DesignStatus { get; set; } = "";
    public string Designer { get; set; } = "";
    public string Engineer { get; set; } = "";
    public string Authority { get; set; } = "";
    public string Mass { get; set; } = "";
    public string SurfaceArea { get; set; } = "";
    public string Volume { get; set; } = "";
    public string SheetMetalRule { get; set; } = "";
    public string FlatPatternWidth { get; set; } = "";
    public string FlatPatternLength { get; set; } = "";
    public string FlatPatternArea { get; set; } = "";
    public string Appearance { get; set; } = "";
    public string Density { get; set; } = "";
    public string LastUpdatedWith { get; set; } = "";

    private Dictionary<string, string> userDefinedProperties = [];

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

    public Dictionary<string, string> UserDefinedProperties
    {
        get => userDefinedProperties;
        set
        {
            userDefinedProperties = value;
            OnPropertyChanged();
        }
    }

    public int Item
    {
        get => item;
        set
        {
            item = value;
            OnPropertyChanged();
        }
    }

    private readonly Dictionary<string, bool> _isExpressionFlags = [];
    private int _expressionStateVersion;
    private bool _suppressExpressionVersion;
    private bool _expressionStateChangedWhileSuppressed;

    public int ExpressionStateVersion
    {
        get => _expressionStateVersion;
        private set
        {
            if (_expressionStateVersion != value)
            {
                _expressionStateVersion = value;
                OnPropertyChanged();
            }
        }
    }

    public void BeginExpressionBatch()
    {
        _suppressExpressionVersion = true;
        _expressionStateChangedWhileSuppressed = false;
    }

    public void EndExpressionBatch()
    {
        _suppressExpressionVersion = false;
        if (_expressionStateChangedWhileSuppressed)
        {
            _expressionStateChangedWhileSuppressed = false;
            ExpressionStateVersion++;
        }
    }

    /// <summary>
    /// Проверяет, является ли указанное свойство выражением
    /// </summary>
    public bool IsPropertyExpression(string propertyName) => 
        _isExpressionFlags.TryGetValue(propertyName, out var isExpression) && isExpression;

    /// <summary>
    /// Устанавливает состояние выражения для свойства
    /// </summary>
    public void SetPropertyExpressionState(string propertyName, bool isExpression)
    {
        var oldValue = IsPropertyExpression(propertyName);
        _isExpressionFlags[propertyName] = isExpression;

        if (oldValue != isExpression)
        {
            if (_suppressExpressionVersion)
            {
                _expressionStateChangedWhileSuppressed = true;
            }
            else
            {
                ExpressionStateVersion++;
            }
        }
    }

    public int Quantity
    {
        get => quantity;
        set
        {
            Debug.WriteLine($"PartData.Quantity setter: PartNumber={PartNumber}, OldValue={quantity}, NewValue={value}, IsOverridden={IsOverridden}");
            if (value > 0 && quantity != value)
            {
                quantity = value;
                IsOverridden = value != OriginalQuantity;
                IsMultiplied = false;
                Debug.WriteLine($"PartData.Quantity setter: Setting IsOverridden={IsOverridden} for {PartNumber}");
                OnPropertyChanged();
            }
            else if (value > 0)
            {
                quantity = value;
                OnPropertyChanged();
            }
        }
    }

    internal void SetQuantityInternal(int value)
    {
        Debug.WriteLine($"PartData.SetQuantityInternal: PartNumber={PartNumber}, OldValue={quantity}, NewValue={value}");
        quantity = value;
        OnPropertyChanged(nameof(Quantity));
    }

    // Реализация интерфейса INotifyPropertyChanged
    public event PropertyChangedEventHandler? PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    // Методы для работы с пользовательскими свойствами
    public void AddUserDefinedProperty(string propertyName, string propertyValue)
    {
        UserDefinedProperties[propertyName] = propertyValue;
        OnPropertyChanged(nameof(UserDefinedProperties));
    }
    public void RemoveUserDefinedProperty(string propertyName) => 
        userDefinedProperties.Remove(propertyName);

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

public class PresetIProperty : INotifyPropertyChanged
{
    public string ColumnHeader { get; set; } = string.Empty; // Заголовок колонки в DataGrid (например, "Обозначение")
    public string ListDisplayName { get; set; } = string.Empty; // Отображение в списке выбора свойств (например, "Обозначение")
    public string InventorPropertyName { get; set; } = string.Empty; // Соответствующее имя свойства iProperty в Inventor (например, "PartNumber")
    public string Category { get; set; } = string.Empty; // Категория свойства для группировки
    public bool IsUserDefined { get; set; } = false; // Является ли это пользовательским свойством

    private bool _isAdded;
    public bool IsAdded
    {
        get => _isAdded;
        set
        {
            if (_isAdded != value)
            {
                _isAdded = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsAdded)));
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

// Структура для отслеживания прогресса сканирования
public class ScanProgress
{
    public int ProcessedItems { get; set; }
    public int TotalItems { get; set; }
    public string CurrentOperation { get; set; } = string.Empty;
    public string CurrentItem { get; set; } = string.Empty;
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
    
    public static UIState CancellingScan => new()
    {
        ScanEnabled = false,
        ScanButtonText = CANCELLING_TEXT,
        ExportEnabled = false,
        ExportButtonText = EXPORT_BUTTON_TEXT,
        ClearEnabled = false,
        ProgressText = "Прерывание...",
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
    
    public static UIState CancellingExport => new()
    {
        ScanEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportEnabled = false,
        ExportButtonText = CANCELLING_TEXT,
        ClearEnabled = false,
        ProgressText = "Прерывание экспорта...",
        ProgressValue = 0,
        InventorUIDisabled = true,
        UpdateScanProgress = false,
        UpdateExportProgress = true
    };
    
    public static UIState PreparingQuickExport => new()
    {
        ScanEnabled = false,
        ScanButtonText = SCAN_BUTTON_TEXT,
        ExportEnabled = true,
        ExportButtonText = CANCEL_BUTTON_TEXT,
        ClearEnabled = false,
        ProgressText = "Сканирование и подготовка данных...",
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
