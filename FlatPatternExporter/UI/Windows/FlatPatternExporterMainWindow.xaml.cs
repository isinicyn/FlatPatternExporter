using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using FlatPatternExporter.Core;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Controls;
using FlatPatternExporter.UI.Models;
using Inventor;
using Binding = System.Windows.Data.Binding;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Size = System.Windows.Size;
using Style = System.Windows.Style;
using TextBox = System.Windows.Controls.TextBox;

namespace FlatPatternExporter.UI.Windows;

public partial class FlatPatternExporterMainWindow : Window, INotifyPropertyChanged
{
    // Inventor API
    private readonly InventorManager _inventorManager = new();
    private Document? _lastScannedDocument;

    // Services
    private readonly LocalizationManager _localizationManager = LocalizationManager.Instance;
    private readonly TokenService _tokenService = new();
    private readonly DocumentScanner _documentScanner;
    private readonly ThumbnailGenerator _thumbnailGenerator = new();
    private readonly DxfExporter _dxfExporter;
    private readonly PartDataReader _partDataReader;
    private readonly ExcelExportService _excelExportService = new();

    // UI elements
    public System.Windows.Controls.Primitives.ToggleButton? ThemeToggleButton { get; private set; }

    // Data and collections
    private readonly ObservableCollection<PartData> _partsData = [];
    private readonly CollectionViewSource _partsDataView;

    // Process state
    private bool _isScanning;
    private bool _isExporting;
    private CancellationTokenSource? _operationCts;
    private int _itemCounter = 1;
    private bool _isUpdatingState;

    // UI state
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
    private string _actualSearchText = string.Empty;
    private readonly DispatcherTimer _searchDelayTimer;

    // Hotkey dictionary for centralized handling
    private readonly Dictionary<Key, Func<Task>> _hotKeyActions;

    // DataGrid column management
    private AdornerLayer? _adornerLayer;
    private HeaderAdorner? _headerAdorner;
    private bool _isColumnDraggedOutside;
    private DataGridColumn? _reorderingColumn;

    // Part filtering settings
    private bool _excludeReferenceParts = true;
    private bool _excludePurchasedParts = true;
    private bool _excludePhantomParts = true;
    private bool _includeLibraryComponents = false;

    // File organization settings
    private bool _organizeByMaterial = false;
    private bool _organizeByThickness = false;

    // File name constructor settings
    private bool _enableFileNameConstructor = false;

    // Optimization settings
    private bool _optimizeDxf = false;

    // DXF export settings
    private bool _enableSplineReplacement = false;
    private SplineReplacementType _selectedSplineReplacement = SplineReplacementType.Lines;
    private AcadVersionType _selectedAcadVersion = AcadVersionType.V2000; // Default: 2000
    private bool _mergeProfilesIntoPolyline = true;
    private bool _rebaseGeometry = true;
    private bool _trimCenterlines = false;

    // Folder and path settings
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

    // Processing method
    private ProcessingMethod _selectedProcessingMethod = ProcessingMethod.BOM;

    // Model state
    private bool _isPrimaryModelState = true;

    // Column templates dictionary (all types)
    private static readonly Dictionary<string, string> ColumnTemplates = InitializeColumnTemplates();

    private static Dictionary<string, string> InitializeColumnTemplates()
    {
        var templates = new Dictionary<string, string>();

        // Automatically add templates for all properties from the centralized registry
        foreach (var definition in PropertyMetadataRegistry.Properties.Values)
        {
            // Use InternalName as key (stable across language changes)
            var columnKey = definition.InternalName;

            // Use unique template if specified
            if (!string.IsNullOrEmpty(definition.ColumnTemplate))
            {
                templates[columnKey] = definition.ColumnTemplate;
            }
            // For editable properties without a unique template, create template with FX indicator
            else if (definition.IsEditable)
            {
                templates[columnKey] = $"{definition.InternalName}WithExpressionTemplate";
            }
        }

        return templates;
    }

    public FlatPatternExporterMainWindow()
    {
        // Initialize services before InitializeComponent to prevent NullReferenceException
        _inventorManager.InitializeInventor();
        _documentScanner = new Core.DocumentScanner(_inventorManager);
        _dxfExporter = new Core.DxfExporter(_inventorManager, _documentScanner, _documentScanner.DocumentCache, _tokenService);
        _partDataReader = new Core.PartDataReader(_inventorManager, _documentScanner, _thumbnailGenerator, Dispatcher);

        InitializeComponent();

        // Initialize theme toggle button from ContentArea StackPanel
        if (TitleBar.ContentArea is StackPanel stackPanel)
        {
            ThemeToggleButton = stackPanel.Children.OfType<System.Windows.Controls.Primitives.ToggleButton>().FirstOrDefault();
        }

        // Initialize hotkey dictionary
        _hotKeyActions = new Dictionary<Key, Func<Task>>
        {
            [Key.F5] = StartScanAsync,
            [Key.F6] = StartNormalExportAsync,
            [Key.F8] = () => { StartClearList(); return Task.CompletedTask; },
            [Key.F9] = StartQuickExportAsync
        };

        UpdateProjectInfo(); // Initialize project folder on startup
        PartsDataGrid.ItemsSource = _partsData;
        PartsDataGrid.PreviewMouseMove += PartsDataGrid_PreviewMouseMove;
        PartsDataGrid.PreviewMouseLeftButtonUp += PartsDataGrid_PreviewMouseLeftButtonUp;
        PartsDataGrid.ColumnReordering += PartsDataGrid_ColumnReordering;
        PartsDataGrid.ColumnReordered += PartsDataGrid_ColumnReordered;

        // Update overlay visibility when column collection changes
        ((INotifyCollectionChanged)PartsDataGrid.Columns).CollectionChanged += (s, e) => UpdateNoColumnsOverlayVisibility();


        // Configure TokenService to work with visual container
        TokenService?.SetTokenContainer(TokenContainer);

        _partsDataView = new CollectionViewSource { Source = _partsData };
        _partsDataView.Filter += PartsData_Filter;

        // Set ItemsSource for DataGrid
        PartsDataGrid.ItemsSource = _partsDataView.View;

        // Configure timer for search delay
        _searchDelayTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1) // 1 second delay
        };
        _searchDelayTimer.Tick += SearchDelayTimer_Tick;

        // Initialize layer settings
        LayerSettings = LayerSettingsHelper.InitializeLayerSettings();
        AvailableColors = LayerSettingsHelper.GetAvailableColors();
        LineTypes = LayerSettingsHelper.GetLineTypes();

        // Initialize ComboBox collections
        InitializeAcadVersions();
        InitializeAvailableTokens();


        // Initialize preset columns based on PropertyMapping
        PresetIProperties = PropertyMetadataRegistry.GetPresetProperties();

        // Set DataContext for current window
        DataContext = this;

        // Add handler for hotkeys F5 (scan), F6 (export), F8 (clear), F9 (quick export)
        KeyDown += MainWindow_KeyDown;

        PresetManager.PropertyChanged += PresetManager_PropertyChanged;
        PresetManager.TemplatePresets.CollectionChanged += TemplatePresets_CollectionChanged;

        // Set selected language in ComboBox
        InitializeLanguageSelection();

        // Subscribe to language change to update button texts and properties
        _localizationManager.LanguageChanged += (s, e) =>
        {
            UpdateButtonTexts();
            UpdatePresetPropertiesLocalization();
        };

        TokenService!.PropertyChanged += TokenService_PropertyChanged;

        SetUIState(UIState.Initial());

        // Initialize overlay visibility
        UpdateNoColumnsOverlayVisibility();

        // Initialize data in TokenService
        _tokenService.UpdatePartsData(_partsData);

    }

    /// <summary>
    /// Applies application settings loaded from App.xaml.cs
    /// </summary>
    public void ApplySettings(ApplicationSettings settings)
    {
        try
        {
            SettingsManager.ApplySettingsToMainWindow(settings, this);
        }
        catch (Exception ex)
        {
            CustomMessageBox.Show(_localizationManager.GetString("Error_SettingsLoad", ex.Message), _localizationManager.GetString("Error_Title"),
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
            CustomMessageBox.Show(_localizationManager.GetString("Error_SettingsSave", ex.Message), _localizationManager.GetString("Error_Title"),
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    public ObservableCollection<LayerSetting> LayerSettings { get; set; }
    public ObservableCollection<LocalizableItem> AvailableColors { get; set; }
    public ObservableCollection<LocalizableItem> LineTypes { get; set; }
    public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }
    public ObservableCollection<AcadVersionItem> AcadVersions { get; set; } = [];
    public ObservableCollection<PropertyMetadataRegistry.PropertyDefinition> AvailableTokens { get; set; } = [];
    public ObservableCollection<PropertyMetadataRegistry.PropertyDefinition> UserDefinedTokens { get; set; } = [];
    public TemplatePresetManager PresetManager { get; } = new();

    // Public properties for CheckBox data binding
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
        set => SetAcadVersion(value, suppressMessage: false);
    }

    internal void SetAcadVersion(AcadVersionType value, bool suppressMessage)
    {
        if (_selectedAcadVersion != value)
        {
            _selectedAcadVersion = value;
            OnPropertyChanged(nameof(SelectedAcadVersion));
            OnPropertyChanged(nameof(IsOptimizeDxfEnabled));

            if (!AcadVersionMapping.SupportsOptimization(_selectedAcadVersion))
            {
                OptimizeDxf = false;
                if (!suppressMessage)
                {
                    var ver = AcadVersionMapping.GetDisplayName(_selectedAcadVersion);
                    CustomMessageBox.Show(
                        _localizationManager.GetString("Info_VersionNotSupported", ver),
                        _localizationManager.GetString("Info_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
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
                OnPropertyChanged(nameof(IsSubfolderCheckBoxEnabled)); // Notify about IsEnabled change
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

    // Computed property for subfolder checkbox IsEnabled
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

    // Computed property for DXF optimization availability
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
    /// Centralized UI state management
    /// </summary>
    private void SetUIState(UIState state)
    {
        ScanButton.IsEnabled = state.ScanEnabled;
        ExportButton.IsEnabled = state.ExportEnabled;
        ExportToExcelButton.IsEnabled = state.ExportEnabled;
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

        _inventorManager.SetInventorUserInterfaceState(state.InventorUIDisabled);
    }

    /// <summary>
    /// Update UI state after operation completion
    /// </summary>
    private void SetUIStateAfterOperation(OperationResult result, OperationType operationType)
    {
        var statusText = result.WasCancelled
            ? _localizationManager.GetString("Status_Interrupted", GetElapsedTime(result.ElapsedTime))
            : operationType == OperationType.Scan
                ? _localizationManager.GetString("Status_PartsFound", result.ProcessedCount, GetElapsedTime(result.ElapsedTime))
                : _localizationManager.GetString("Status_Completed", GetElapsedTime(result.ElapsedTime));
        
        var state = UIState.CreateAfterOperationState(_partsData.Count > 0, result.WasCancelled, statusText);
        SetUIState(state);
    }

    /// <summary>
    /// Centralized display of operation results to user
    /// </summary>
    private void ShowOperationResult(OperationResult result, OperationType operationType, bool isQuickMode = false)
    {
        if (result.WasCancelled)
        {
            var operationName = operationType == OperationType.Scan
                ? _localizationManager.GetString("Operation_Scanning")
                : _localizationManager.GetString("Operation_Export");
            CustomMessageBox.Show(_localizationManager.GetString("Info_OperationInterrupted", operationName),
                _localizationManager.GetString("Info_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        switch (operationType)
        {
            case OperationType.Scan:
                // For scanning, show special messages about conflicts and references
                if (_documentScanner.ConflictAnalyzer.ConflictCount > 0)
                {
                    CustomMessageBox.Show(_localizationManager.GetString("Info_ConflictsDetected", _documentScanner.ConflictAnalyzer.ConflictCount),
                        _localizationManager.GetString("Info_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else if (_documentScanner.HasMissingReferences)
                {
                    var messageKey = result.ProcessingMethod == ProcessingMethod.BOM
                        ? "Info_BrokenReferences_BOM"
                        : "Info_BrokenReferences_Traverse";
                    CustomMessageBox.Show(_localizationManager.GetString(messageKey),
                        _localizationManager.GetString("Info_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                break;
                
            case OperationType.Export:
                var exportTitle = isQuickMode
                    ? _localizationManager.GetString("Info_QuickExportCompleted")
                    : _localizationManager.GetString("Info_ExportCompleted");
                CustomMessageBox.Show(
                    this,
                    _localizationManager.GetString("Info_ExportStatistics",
                        result.ProcessedCount + result.SkippedCount, result.SkippedCount,
                        result.ProcessedCount, GetElapsedTime(result.ElapsedTime)),
                    exportTitle, MessageBoxButton.OK, MessageBoxImage.Information);
                break;
        }
    }

    /// <summary>
    /// Centralized error handling for operations
    /// </summary>
    private static async Task<T> ExecuteWithErrorHandlingAsync<T>(
        Func<Task<T>> operation,
        string operationName) where T : OperationResult, new()
    {
        try
        {
            return await operation();
        }
        catch (OperationCanceledException)
        {
            // Operation was cancelled - return result with cancellation flag
            var result = new T
            {
                WasCancelled = true
            };
            return result;
        }
        catch (Exception ex)
        {
            var result = new T();
            result.Errors.Add(LocalizationManager.Instance.GetString("Error_Operation", operationName, ex.Message));
            return result;
        }
    }


    /// <summary>
    /// Centralized validation of active document with error display
    /// </summary>
    private DocumentValidationResult? ValidateDocumentOrShowError()
    {
        var validation = _inventorManager.ValidateActiveDocument();
        if (!validation.IsValid)
        {
            CustomMessageBox.Show(validation.ErrorMessage, _localizationManager.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
            return null;
        }
        return validation;
    }

    /// <summary>
    /// Centralized export context preparation with error display
    /// </summary>
    private async Task<ExportContext?> PrepareExportContextOrShowError(Document document, bool requireScan = true, bool showProgress = false)
    {
        var exportOptions = CreateExportOptions();
        var context = await _dxfExporter.PrepareExportContextAsync(document, requireScan, showProgress, _lastScannedDocument, exportOptions);
        if (!context.IsValid)
        {
            CustomMessageBox.Show(context.ErrorMessage, _localizationManager.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
            if (context.ErrorMessage.Contains(LocalizationManager.Instance.GetString("Operation_ScanningKey")))
                MultiplierTextBox.Text = "1";
            return null;
        }
        return context;
    }

    private ExportOptions CreateExportOptions()
    {
        return new ExportOptions
        {
            SelectedExportFolder = SelectedExportFolder,
            FixedFolderPath = FixedFolderPath,
            EnableSubfolder = EnableSubfolder,
            SubfolderName = SubfolderNameTextBox.Text,
            Multiplier = int.TryParse(MultiplierTextBox.Text, out var m) ? m : 1,
            SelectedProcessingMethod = SelectedProcessingMethod,
            ExcludeReferenceParts = ExcludeReferenceParts,
            ExcludePurchasedParts = ExcludePurchasedParts,
            ExcludePhantomParts = ExcludePhantomParts,
            IncludeLibraryComponents = IncludeLibraryComponents,
            OrganizeByMaterial = OrganizeByMaterial,
            OrganizeByThickness = OrganizeByThickness,
            EnableFileNameConstructor = EnableFileNameConstructor,
            OptimizeDxf = OptimizeDxf,
            EnableSplineReplacement = EnableSplineReplacement,
            SelectedSplineReplacement = SelectedSplineReplacement,
            SplineTolerance = SplineToleranceTextBox.Text,
            SelectedAcadVersion = SelectedAcadVersion,
            MergeProfilesIntoPolyline = MergeProfilesIntoPolyline,
            RebaseGeometry = RebaseGeometry,
            TrimCenterlines = TrimCenterlines,
            LayerSettings = LayerSettings.ToList(),
            ShowFileLockedDialogs = true
        };
    }

    /// <summary>
    /// Centralized operation initialization
    /// </summary>
    private void InitializeOperation(UIState operationState, ref bool operationFlag)
    {
        SetUIState(operationState);
        operationFlag = true;
        _operationCts = new CancellationTokenSource();
    }

    /// <summary>
    /// Centralized operation completion
    /// </summary>
    private void CompleteOperation(OperationResult result, OperationType operationType, ref bool operationFlag, bool isQuickMode = false)
    {
        operationFlag = false;
        SetUIStateAfterOperation(result, operationType);
        ShowOperationResult(result, operationType, isQuickMode);
    }

    /// <summary>
    /// Centralized creation of export operation result
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
    /// Centralized handling of export cancellation
    /// </summary>
    private bool HandleExportCancellation()
    {
        if (_isExporting)
        {
            _operationCts?.Cancel();
            SetUIState(UIState.CancellingExport());
            return true;
        }
        return false;
    }

    protected override void OnClosed(EventArgs e)
    {
        // Cancel active operations and free resources
        _operationCts?.Cancel();
        _operationCts?.Dispose();

        // Save settings on close
        SaveSettings();

        base.OnClosed(e);

        // Close all child windows
        foreach (Window window in System.Windows.Application.Current.Windows)
            if (window != this)
                window.Close();

        // If additional threads were created, ensure they are terminated
    }
    private void AddPresetIPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        var selectIPropertyWindow = new SelectIPropertyWindow(
            PresetIProperties,
            inventorPropertyName => PartsDataGrid.Columns.Any(c => c.SortMemberPath == inventorPropertyName),
            AddIPropertyColumn,
            AddUserDefinedIPropertyColumn,
            RemoveDataGridColumn);
        selectIPropertyWindow.ShowDialog();
    }

    private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv",
                DefaultExt = ".xlsx",
                FileName = $"Export_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            var isXlsx = System.IO.Path.GetExtension(saveFileDialog.FileName).ToLower() == ".xlsx";

            if (isXlsx)
            {
                _excelExportService.ExportToExcel(saveFileDialog.FileName, PartsDataGrid, _partsDataView.View);
            }
            else
            {
                _excelExportService.ExportToCsv(saveFileDialog.FileName, PartsDataGrid, _partsDataView.View);
            }

            CustomMessageBox.Show(
                this,
                _localizationManager.GetString("Message_ExportSuccess", saveFileDialog.FileName),
                _localizationManager.GetString("Title_Success"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            CustomMessageBox.Show(
                this,
                _localizationManager.GetString("Message_ExportError", ex.Message),
                _localizationManager.GetString("Title_Error"),
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void MainWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        // Get the original element
        var clickedElement = e.OriginalSource as DependencyObject;

        // Move up the visual element tree
        while (clickedElement != null && !(clickedElement is Visual || clickedElement is Visual3D))
        {
            // Check that it's not a Run element
            if (clickedElement is Run)
                // Simply exit if we encounter a Run
                return;
            clickedElement = VisualTreeHelper.GetParent(clickedElement);
        }

        // Get event source
        var source = e.OriginalSource as DependencyObject;

        // Search for DataGrid in the visual element hierarchy
        while (source != null && source is not DataGrid) source = VisualTreeHelper.GetParent(source);

        // If DataGrid not found (click was outside DataGrid), clear selection
        if (source == null) PartsDataGrid.UnselectAll();
    }

    private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
        // Clear text field
        SearchTextBox.Text = string.Empty;
        ClearSearchButton.IsEnabled = false;
        SearchTextBox.Focus(); // Set focus to search field
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _actualSearchText = SearchTextBox.Text.Trim().ToLower();

        // Enable or disable clear button depending on text presence
        ClearSearchButton.IsEnabled = !string.IsNullOrEmpty(_actualSearchText);

        // Restart timer on text change
        _searchDelayTimer.Stop();
        _searchDelayTimer.Start();
    }

    private void AddCustomTextButton_Click(object sender, RoutedEventArgs e)
    {
        var customText = CustomTextBox.Text;
        if (string.IsNullOrEmpty(customText)) return;

        // Add text via TokenService
        TokenService?.AddCustomText(customText);

        // Clear input field
        CustomTextBox.Text = string.Empty;
    }

    private void AddSymbolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is string symbol)
        {
            // Add symbol via TokenService
            TokenService?.AddCustomText(symbol);
        }
    }

    private void CustomTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddCustomTextButton_Click(sender, new RoutedEventArgs());
            e.Handled = true;
        }
    }
    private void SearchDelayTimer_Tick(object? sender, EventArgs e)
    {
        _searchDelayTimer.Stop();
        _partsDataView.View.Refresh(); // Update filtering
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

        // Iterate through all visible DataGrid columns
        foreach (var column in PartsDataGrid.Columns)
        {
            var columnHeader = column.Header.ToString();
            if (string.IsNullOrEmpty(columnHeader)) continue;

            // Search by standard properties
            var metadata = PropertyMetadataRegistry.Properties.Values
                .FirstOrDefault(p => p.ColumnHeader == columnHeader);

            if (metadata != null && !metadata.IsSearchable) continue;

            string? searchValue = null;

            // Get value from standard property
            if (metadata != null)
            {
                var propInfo = typeof(PartData).GetProperty(metadata.InternalName);
                searchValue = propInfo?.GetValue(partData)?.ToString();
            }
            // Get value from user defined property
            // Find User Defined Property by ColumnHeader
            var userProp = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeader);
            if (userProp != null && partData.UserDefinedProperties.TryGetValue(userProp.InventorPropertyName!, out string? value))
            {
                searchValue = value;
            }

            // Check match
            if (!string.IsNullOrEmpty(searchValue) &&
                searchValue.Contains(_actualSearchText, StringComparison.CurrentCultureIgnoreCase))
            {
                e.Accepted = true;
                return;
            }
        }

        e.Accepted = false;
    }

    public void AddIPropertyColumn(PresetIProperty iProperty)
    {
        // Check if column with this InventorPropertyName already exists
        if (PartsDataGrid.Columns.Any(c => c.SortMemberPath == iProperty.InventorPropertyName))
            return;

        // Get property metadata to check sortability
        var isSortable = true;
        if (PropertyMetadataRegistry.Properties.TryGetValue(iProperty.InventorPropertyName, out var propertyDef))
        {
            isSortable = propertyDef.IsSortable;
        }

        DataGridColumn column;

        // Check columns with templates (use InventorPropertyName as stable key)
        if (ColumnTemplates.TryGetValue(iProperty.InventorPropertyName, out var templateName))
        {
            DataTemplate template;

            // If this is an editable property with fx indicator - create dynamically
            if (templateName.EndsWith("WithExpressionTemplate"))
            {
                template = FindResource("EditableWithFxTemplate") as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException("EditableWithFxTemplate not found", "EditableWithFxTemplate");
            }
            else
            {
                // For other templates, search in resources
                template = FindResource(templateName) as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException($"Template {templateName} not found", templateName);
            }

            var templateColumn = new DataGridTemplateColumn
            {
                Header = iProperty.ColumnHeader,
                CellTemplate = template,
                SortMemberPath = iProperty.InventorPropertyName,
                CanUserSort = isSortable,
                IsReadOnly = !templateName.StartsWith("Editable"),
                ClipboardContentBinding = new Binding(iProperty.InventorPropertyName)
            };

            // Pass property path via cell Tag for universal template
            if (templateName.EndsWith("WithExpressionTemplate"))
            {
                templateColumn.CellStyle = CreateCellTagStyle(iProperty.InventorPropertyName);
            }

            column = templateColumn;
        }
        else
        {
            // Create regular text column via centralized method
            column = CreateTextColumn(iProperty.ColumnHeader, iProperty.InventorPropertyName, isSortable);
        }

        // Add column to the end
        PartsDataGrid.Columns.Add(column);

        // Update property states in selection window
        var selectIPropertyWindow =
            System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
        selectIPropertyWindow?.UpdatePropertyStates();

        // Fill data for new column
        FillPropertyData(iProperty.InventorPropertyName);
    }

    private void PartsDataGrid_ColumnReordering(object? sender, DataGridColumnReorderingEventArgs e)
    {
        _reorderingColumn = e.Column;
        _isColumnDraggedOutside = false;

        // Create phantom header using style from XAML
        var header = new TextBlock
        {
            Text = _reorderingColumn.Header.ToString(),
            Width = _reorderingColumn.ActualWidth,
            Style = (Style)FindResource("PhantomColumnHeaderStyle")
        };

        // Create and add Adorner
        _adornerLayer = AdornerLayer.GetAdornerLayer(PartsDataGrid);
        _headerAdorner = new HeaderAdorner(PartsDataGrid, header);
        _adornerLayer.Add(_headerAdorner);

        // Initially place Adorner where the header is located
        UpdateAdornerPosition(e);
    }

    private void PartsDataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (_headerAdorner == null || _reorderingColumn == null)
            return;

        // Update Adorner position
        UpdateAdornerPosition(null);

        // Get current mouse position relative to DataGrid
        var position = e.GetPosition(PartsDataGrid);

        // Check if mouse left DataGrid bounds
        if (position.X > PartsDataGrid.ActualWidth || position.Y < 0 || position.Y > PartsDataGrid.ActualHeight)
        {
            _isColumnDraggedOutside = true;
            SetAdornerDeleteMode(true); // Switch to delete mode
        }
        else
        {
            _isColumnDraggedOutside = false;
            SetAdornerDeleteMode(false); // Return to normal mode
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

            // Remove fixed sizes and set auto-sizing
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

            // Use centralized method with removeDataOnly=true for Adorner
            RemoveDataGridColumn(columnName, removeDataOnly: true);
        }

        _reorderingColumn = null;
        _isColumnDraggedOutside = false;
    }

    private void PartsDataGrid_ColumnReordered(object? sender, DataGridColumnEventArgs e)
    {
        _reorderingColumn = null;
        _isColumnDraggedOutside = false;

        // Remove Adorner
        if (_adornerLayer != null && _headerAdorner != null)
        {
            _adornerLayer.Remove(_headerAdorner);
            _headerAdorner = null;
        }
    }

    private void UpdateAdornerPosition(DataGridColumnReorderingEventArgs? _)
    {
        if (_headerAdorner == null)
            return;

        // Get current mouse position relative to DataGrid parent
        var position = Mouse.GetPosition(_adornerLayer);

        // Apply transformation to move Adorner
        _headerAdorner.RenderTransform = new TranslateTransform(
            position.X - _headerAdorner.RenderSize.Width / 2,
            position.Y - _headerAdorner.RenderSize.Height / 1.8
        );
    }

    private async void MainWindow_KeyDown(object sender, KeyEventArgs e)
    {
        // Check if there's a handler for the pressed key
        if (_hotKeyActions.TryGetValue(e.Key, out var action))
        {
            await action();
            e.Handled = true;
        }
    }

    private async Task StartScanAsync()
    {
        // Check that no other operations are running
        if (_isExporting || _isScanning)
            return;

        // Call scanning logic from ScanButton_Click
        await PerformScanOperationAsync();
    }

    private async Task PerformScanOperationAsync()
    {
        // Document validation
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Configure UI for scanning
        InitializeOperation(UIState.Scanning(), ref _isScanning);
        // Missing references flag reset is performed inside ScanService

        // Progress for assembly structure scanning
        var scanProgress = new Progress<ScanProgress>(progress =>
        {
            var progressState = UIState.Scanning();
            progressState.ProgressValue = progress.TotalItems > 0 ? (double)progress.ProcessedItems / progress.TotalItems * 100 : 0;
            progressState.ProgressText = $"{progress.CurrentOperation} - {progress.CurrentItem}";
            SetUIState(progressState);
        });

        // Execute scanning
        var result = await ExecuteWithErrorHandlingAsync(
            () => ScanDocumentAsync(validation.Document!, updateUI: true, scanProgress, _operationCts!.Token),
            LocalizationManager.Instance.GetString("Operation_Scanning"));

        // Complete operation
        CompleteOperation(result, OperationType.Scan, ref _isScanning);
    }

    private async Task StartNormalExportAsync()
    {
        // Check that no other operations are running
        if (_isExporting || _isScanning)
            return;

        // Check that there's data to export (like button does)
        if (_partsData.Count == 0)
        {
            CustomMessageBox.Show(_localizationManager.GetString("Info_NoDataForExport"),
                _localizationManager.GetString("Info_Title"),
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Perform normal export
        await PerformNormalExportAsync();
    }

    private async Task PerformNormalExportAsync()
    {
        // Document validation
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Prepare export context (requires prior scanning)
        var context = await PrepareExportContextOrShowError(validation.Document!, requireScan: true, showProgress: true);
        if (context == null) return;

        // Configure UI for export
        InitializeOperation(UIState.Exporting(), ref _isExporting);
        var stopwatch = Stopwatch.StartNew();

        // Execute export via centralized error handling
        var partsDataList = _partsData.Where(p => context.SheetMetalParts.ContainsKey(p.PartNumber)).ToList();
        var exportOptions = CreateExportOptions();
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            var exportProgress = new Progress<double>(UpdateExportProgress);
            await Task.Run(() => _dxfExporter.ExportDXF(partsDataList, context.TargetDirectory, context.Multiplier,
                exportOptions, ref processedCount, ref skippedCount, context.GenerateThumbnails, exportProgress, _operationCts!.Token), _operationCts!.Token);

            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, LocalizationManager.Instance.GetString("Operation_Export"));

        // Complete operation
        CompleteOperation(result, OperationType.Export, ref _isExporting);
    }

    private void StartClearList()
    {
        // Check that no export or scanning operations are running
        if (_isExporting || _isScanning)
        {
            CustomMessageBox.Show(_localizationManager.GetString("Info_CannotClearDuringOperation"),
                _localizationManager.GetString("Info_Title"),
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Perform clearing via existing method
        ClearList_Click(this, null!);
    }

    private async Task StartQuickExportAsync()
    {
        // Check that no other operations are running
        if (_isExporting || _isScanning)
            return;

        await ExportWithoutScan();
    }

    /// <summary>
    /// Update export progress via centralized UI system
    /// </summary>
    private void UpdateExportProgress(double progressValue)
    {
        var progressState = UIState.Exporting();
        progressState.ProgressValue = progressValue;
        progressState.UpdateExportProgress = true;
        progressState.UpdateScanProgress = false;
        SetUIState(progressState);
    }

    private void MultiplierTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        // Check current text together with new input
        if (sender is not TextBox textBox) return;
        var newText = textBox.Text.Insert(textBox.CaretIndex, e.Text);

        // Check if text is a positive number greater than zero
        e.Handled = !IsTextAllowed(newText);
    }

    private static bool IsTextAllowed(string text)
    {
        // Regular expression for checking positive numbers
        return PositiveNumberRegex().IsMatch(text);
    }

    [GeneratedRegex("^[1-9][0-9]*$")]
    private static partial Regex PositiveNumberRegex();

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
    /// Unified document scanning method
    /// </summary>
    private async Task<OperationResult> ScanDocumentAsync(
        Document document,
        bool updateUI = true,
        IProgress<ScanProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        var result = new OperationResult();
        
        try
        {
            if (updateUI)
            {
                // Get and display document information
                UpdateDocumentInfo(document);
                _partsData.Clear();
                _itemCounter = 1;
            }

            ClearConflictData();

            var sheetMetalParts = new Dictionary<string, int>();

            // Create scanning options
            var scanOptions = new Core.ScanOptions
            {
                ExcludeReferenceParts = ExcludeReferenceParts,
                ExcludePurchasedParts = ExcludePurchasedParts,
                ExcludePhantomParts = ExcludePhantomParts,
                IncludeLibraryComponents = IncludeLibraryComponents
            };

            // Use ScanService for scanning
            var scanResult = await _documentScanner.ScanDocumentAsync(
                document,
                SelectedProcessingMethod,
                scanOptions,
                progress,
                cancellationToken);

            sheetMetalParts = scanResult.SheetMetalParts;
            result.ProcessedCount = scanResult.ProcessedCount;
            result.ProcessingMethod = scanResult.ProcessingMethod;

            // Process parts for UI
            if (updateUI && sheetMetalParts.Count > 0)
            {
                // Conflict analysis already performed in ScanService
                if (_documentScanner.ConflictAnalyzer.HasConflicts)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        ConflictFilesButton.IsEnabled = true;
                    });
                }

                // Switch progress to part processing
                await Dispatcher.InvokeAsync(() =>
                {
                    var processingState = UIState.Scanning();
                    processingState.ProgressText = _localizationManager.GetString("Status_ProcessingParts");
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

                        var partData = await _partDataReader.GetPartDataAsync(part.Key, part.Value, itemCounter++);
                        if (partData != null)
                        {
                            ((IProgress<PartData>)partProgress).Report(partData);
                        }

                        processedParts++;
                        await Dispatcher.InvokeAsync(() =>
                        {
                            var progressState = UIState.Scanning();
                            progressState.ProgressValue = totalParts > 0 ? (double)processedParts / totalParts * 100 : 0;
                            progressState.ProgressText = _localizationManager.GetString("Status_ProcessingPartsProgress", processedParts, totalParts);
                            SetUIState(progressState);
                        });
                    }
                }, cancellationToken);
            }

            if (updateUI && int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
                _partDataReader.UpdateQuantitiesWithMultiplier(_partsData, multiplier);

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
        // If scanning is already in progress, perform cancellation
        if (_isScanning)
        {
            _operationCts?.Cancel();
            SetUIState(UIState.CancellingScan());
            return;
        }

        // Perform scanning
        await PerformScanOperationAsync();
    }
    private void ConflictFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_documentScanner.ConflictAnalyzer.ConflictFileDetails == null || _documentScanner.ConflictAnalyzer.ConflictFileDetails.Count == 0)
        {
            return;
        }

        // Create and show window with conflict details
        var conflictWindow = new ConflictDetailsWindow(_documentScanner.ConflictAnalyzer.ConflictFileDetails, _inventorManager.OpenInventorDocument)
        {
            Owner = this
        };
        conflictWindow.ShowDialog();
    }


    private void UpdateSubfolderOptionsFromProperty()
    {
        // "Next to part" option - checkbox should be disabled
        if (SelectedExportFolder == ExportFolderType.PartFolder)
        {
            EnableSubfolder = false;
        }

        // Special logic for ProjectFolder
        if (SelectedExportFolder == ExportFolderType.ProjectFolder)
        {
            UpdateProjectInfo(); // Set project information on selection
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
                    caretIndex--; // Decrease cursor position by 1
                }
        }

        // Return cursor to correct position after character removal
        if (textBox != null)
            textBox.CaretIndex = Math.Max(caretIndex, 0);
    }
    // Method to update project information in UI
    private void UpdateProjectInfo()
    {
        _inventorManager.SetProjectFolderInfo();
        ProjectNameRun.Text = _inventorManager.ProjectName;
        ProjectStackPanelItem.ToolTip = _inventorManager.ProjectWorkspacePath;
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

    private void InitializeAvailableTokens()
    {
        AvailableTokens = [];
        UserDefinedTokens = [];
        
        RefreshAvailableTokens();
        
        // Subscribe to changes in user-defined properties collection
        PropertyMetadataRegistry.UserDefinedProperties.CollectionChanged += (s, e) =>
        {
            RefreshAvailableTokens();
        };
    }

    private void RefreshAvailableTokens()
    {
        AvailableTokens.Clear();
        UserDefinedTokens.Clear();

        // Get all tokenizable properties from centralized registry
        var tokenizableProperties = PropertyMetadataRegistry.GetTokenizableProperties();

        // Separate into standard and user-defined properties
        var standardProperties = tokenizableProperties.Where(p => p.Type != PropertyMetadataRegistry.PropertyType.UserDefined).OrderBy(p => p.DisplayName);
        var userDefinedProperties = tokenizableProperties.Where(p => p.Type == PropertyMetadataRegistry.PropertyType.UserDefined).OrderBy(p => p.DisplayName);

        // Add standard properties
        foreach (var property in standardProperties)
        {
            AvailableTokens.Add(property);
        }

        // Add user-defined properties
        foreach (var property in userDefinedProperties)
        {
            UserDefinedTokens.Add(property);
        }
    }

    /// <summary>
    /// Clears all collections related to conflicts for operation isolation
    /// </summary>
    private void ClearConflictData()
    {
        _documentScanner.ConflictAnalyzer.Clear();
        ConflictFilesButton.IsEnabled = false; // Disable button when clearing data
    }

    private void TokenListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        var listBox = sender as System.Windows.Controls.ListBox;
        if (listBox?.SelectedItem is PropertyMetadataRegistry.PropertyDefinition selectedProperty)
        {
            TokenService?.AddToken($"{{{selectedProperty.TokenName}}}");
        }
    }

    private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            FixedFolderPath = dialog.SelectedPath;
        }
    }

    private async Task ExportWithoutScan()
    {
        // Clear table before starting hidden export
        ClearList_Click(this, null!);

        // Document validation
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Configure UI for quick export preparation
        InitializeOperation(UIState.PreparingQuickExport(), ref _isExporting);

        // Prepare export context (without requiring prior scanning and without showing progress)
        var context = await PrepareExportContextOrShowError(validation.Document!, requireScan: false, showProgress: false);
        if (context == null)
        {
            // Restore state on error
            SetUIState(UIState.CreateAfterOperationState(false, false, _localizationManager.GetString("Error_ExportPreparation")));
            _isExporting = false;
            return;
        }

        // Create temporary PartData list for export with full data (this is also a lengthy operation)
        var tempPartsDataList = new List<PartData>();
        var itemCounter = 1;
        var totalParts = context.SheetMetalParts.Count;

        foreach (var part in context.SheetMetalParts)
        {
            // Update progress every 5 parts or at important moments
            if (itemCounter % 5 == 1 || itemCounter == totalParts)
            {
                var progressState = UIState.PreparingQuickExport();
                progressState.ProgressText = _localizationManager.GetString("Status_PreparingPartData", itemCounter, totalParts);
                SetUIState(progressState);
                await Task.Delay(1); // Minimal delay for UI update
            }

            var partData = await _partDataReader.GetPartDataAsync(part.Key, part.Value * context.Multiplier, itemCounter++, loadThumbnail: false);
            if (partData != null)
            {
                // Don't call SetQuantityInternal again - quantity already set correctly in GetPartDataAsync
                partData.IsMultiplied = context.Multiplier > 1;
                tempPartsDataList.Add(partData);
            }
        }

        // Switch to export state when file export ACTUALLY begins
        SetUIState(UIState.Exporting());

        var stopwatch = Stopwatch.StartNew();

        // Execute export via centralized error handling
        context.GenerateThumbnails = false;
        var exportOptions = CreateExportOptions();
        exportOptions.ShowFileLockedDialogs = true;
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            var exportProgress = new Progress<double>(UpdateExportProgress);
            await Task.Run(() => _dxfExporter.ExportDXF(tempPartsDataList, context.TargetDirectory, context.Multiplier,
                exportOptions, ref processedCount, ref skippedCount, context.GenerateThumbnails, exportProgress, _operationCts!.Token), _operationCts!.Token);

            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, "quick export");

        // Complete operation
        CompleteOperation(result, OperationType.Export, ref _isExporting, isQuickMode: true);
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        // If export is already in progress, perform cancellation
        if (HandleExportCancellation()) return;

        // Perform normal export
        await PerformNormalExportAsync();
    }

    private static string GetElapsedTime(TimeSpan timeSpan)
    {
        var minutesShort = LocalizationManager.Instance.GetString("Time_Minutes_Short");
        var secondsShort = LocalizationManager.Instance.GetString("Time_Seconds_Short");

        if (timeSpan.TotalMinutes >= 1)
            return $"{(int)timeSpan.TotalMinutes} {minutesShort} {timeSpan.Seconds}.{timeSpan.Milliseconds:D3} {secondsShort}";
        return $"{timeSpan.Seconds}.{timeSpan.Milliseconds:D3} {secondsShort}";
    }

    private async void ExportSelectedDXF_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();

        var itemsWithoutFlatPattern = selectedItems.Where(p => !p.HasFlatPattern).ToList();
        if (itemsWithoutFlatPattern.Count == selectedItems.Count)
        {
            CustomMessageBox.Show(_localizationManager.GetString("Message_NoFlatPatternsSelected"), _localizationManager.GetString("MessageBox_Information"), MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (itemsWithoutFlatPattern.Count > 0)
        {
            var dialogResult =
                CustomMessageBox.Show(_localizationManager.GetString("Message_SomeFlatPatternsSkipped"),
                    _localizationManager.GetString("MessageBox_Information"), MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (dialogResult == MessageBoxResult.No) return;

            selectedItems = [.. selectedItems.Except(itemsWithoutFlatPattern)];
        }

        // Document validation
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Prepare export context
        var context = await PrepareExportContextOrShowError(validation.Document!, requireScan: true, showProgress: false);
        if (context == null) return;

        // Configure UI for export
        InitializeOperation(UIState.Exporting(), ref _isExporting);
        var stopwatch = Stopwatch.StartNew();

        // Execute export via centralized error handling
        var exportOptions = CreateExportOptions();
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = itemsWithoutFlatPattern.Count;
            var exportProgress = new Progress<double>(UpdateExportProgress);
            await Task.Run(() => _dxfExporter.ExportDXF(selectedItems, context.TargetDirectory, context.Multiplier,
                exportOptions, ref processedCount, ref skippedCount, context.GenerateThumbnails, exportProgress, _operationCts!.Token), _operationCts!.Token);

            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, "selected parts export");

        // Complete operation
        CompleteOperation(result, OperationType.Export, ref _isExporting);
    }

    private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
        {
            _partDataReader.UpdateQuantitiesWithMultiplier(_partsData, multiplier);

            // Check for null before changing button state
            if (ClearMultiplierButton != null)
                ClearMultiplierButton.IsEnabled = multiplier > 1;
        }
        else
        {
            // If entered value is incorrect, reset text to "1" and disable reset button
            MultiplierTextBox.Text = "1";
            _partDataReader.UpdateQuantitiesWithMultiplier(_partsData, 1);

            if (ClearMultiplierButton != null) ClearMultiplierButton.IsEnabled = false;
        }
    }

    private void ClearMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        // Reset multiplier to 1
        MultiplierTextBox.Text = "1";
        _partDataReader.UpdateQuantitiesWithMultiplier(_partsData, 1);

        // Disable reset button
        ClearMultiplierButton.IsEnabled = false;
    }

    private void UpdateDocumentInfo(Document? doc)
    {
        var docInfo = _partDataReader.GetDocumentInfo(doc);

        if (string.IsNullOrEmpty(docInfo.DocumentType))
        {
            var noDocState = UIState.Initial();
            noDocState.ProgressText = _localizationManager.GetString("Status_NoDocumentInfo");
            SetUIState(noDocState);
            DocumentTypeLabel.Text = "";
            PartNumberLabel.Text = "";
            DescriptionLabel.Text = "";
            ModelStateLabel.Text = "";
            IsPrimaryModelState = true; // Reset model state to default
            return;
        }

        // Fill individual fields in document information block
        DocumentTypeLabel.Text = docInfo.DocumentType;
        PartNumberLabel.Text = docInfo.PartNumber;
        DescriptionLabel.Text = docInfo.Description;
        ModelStateLabel.Text = docInfo.ModelState;

        // Set property for style trigger
        IsPrimaryModelState = docInfo.IsPrimaryModelState;
    }

    private void ClearList_Click(object sender, RoutedEventArgs e)
    {
        // Clear data in table or any other data source
        _partsData.Clear();
        SetUIState(UIState.CreateClearedState());

        // Reset document information
        UpdateDocumentInfo(null);

        // Update TokenService to display placeholders
        _tokenService.UpdatePartsData(_partsData);

        // Disable "Conflicts" button and clear conflicts list
        _documentScanner.ConflictAnalyzer.Clear();
        ConflictFilesButton.IsEnabled = false;

        // Clear document cache
        _documentScanner.DocumentCache.ClearCache();
    }

    private void RemoveSelectedRows_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();

        var result = CustomMessageBox.Show(_localizationManager.GetString("Message_ConfirmDelete", selectedItems.Count),
            _localizationManager.GetString("MessageBox_Confirmation"), MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            foreach (var item in selectedItems)
            {
                _partsData.Remove(item);
            }

            if (_partsData.Count == 0)
            {
                ClearList_Click(sender, e);
            }
            else
            {
                // Update TokenService after removing rows
                _tokenService.UpdatePartsData(_partsData);

                var deletionState = UIState.CreateAfterOperationState(_partsData.Count > 0, false, $"Removed {selectedItems.Count} row(s)");
                SetUIState(deletionState);
            }
        }
    }

    private void PartsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_partsData.Count == 0)
        {
            e.Handled = true; // Prevent opening context menu if table is empty
            return;
        }

        // Manage menu item availability based on selection
        var hasSelection = PartsDataGrid.SelectedItems.Count > 0;
        var hasSingleSelection = PartsDataGrid.SelectedItems.Count == 1;
        var hasOverriddenSelection = hasSelection && PartsDataGrid.SelectedItems.Cast<PartData>().Any(item => item.IsOverridden);

        ExportSelectedDXFMenuItem.IsEnabled = hasSelection;
        OpenFileLocationMenuItem.IsEnabled = hasSelection;
        OpenSelectedModelsMenuItem.IsEnabled = hasSelection;
        ResetQuantityMenuItem.IsEnabled = hasOverriddenSelection;
        RemoveSelectedRowsMenuItem.IsEnabled = hasSelection;
    }

    private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();

        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = _documentScanner.DocumentCache.GetCachedPartPath(partNumber) ?? _inventorManager.GetPartDocumentFullPath(partNumber);

            // Check for null before using fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                CustomMessageBox.Show(_localizationManager.GetString("Error_FileNotFoundForPart", partNumber), _localizationManager.GetString("MessageBox_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Check file existence
            if (System.IO.File.Exists(fullPath))
                try
                {
                    var argument = "/select, \"" + fullPath + "\"";
                    Process.Start("explorer.exe", argument);
                }
                catch (Exception ex)
                {
                    CustomMessageBox.Show(_localizationManager.GetString("Error_ExplorerOpen", fullPath, ex.Message),
                        _localizationManager.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            else
                CustomMessageBox.Show(_localizationManager.GetString("Error_FileNotFound", fullPath),
                    _localizationManager.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSelectedModels_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();

        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = _documentScanner.DocumentCache.GetCachedPartPath(partNumber) ?? _inventorManager.GetPartDocumentFullPath(partNumber);
            var targetModelState = item.ModelState;

            // Check for null before using fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                CustomMessageBox.Show(_localizationManager.GetString("Error_FileNotFoundForPart", partNumber), _localizationManager.GetString("MessageBox_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Open file with specified model state
            _inventorManager.OpenInventorDocument(fullPath, targetModelState);
        }
    }

    private void PartsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Update preview with selected row
        var selectedPart = PartsDataGrid.SelectedItem as PartData;
        _tokenService.UpdatePreviewWithSelectedData(selectedPart);
    }

    public void FillPropertyData(string propertyName)
    {
        _partDataReader.FillPropertyData(_partsData, propertyName);
    }

    private void RemoveUserDefinedIPropertyColumn(string internalName)
    {
        // Get property name from InternalName (UDP_PropertyName -> PropertyName)
        var propertyName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);

        // Remove all data related to this User Defined Property
        foreach (var partData in _partsData)
        {
            if (partData.UserDefinedProperties.ContainsKey(propertyName))
                partData.RemoveUserDefinedProperty(propertyName);
        }

        // Remove from registry
        PropertyMetadataRegistry.RemoveUserDefinedProperty(internalName);
    }

    private void RemoveUserDefinedIPropertyData(string internalName)
    {
        // Get property name from InternalName (UDP_PropertyName -> PropertyName)
        var propertyName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);

        // Remove only UDP data from PartData (without removing from registry)
        foreach (var partData in _partsData)
        {
            if (partData.UserDefinedProperties.ContainsKey(propertyName))
                partData.RemoveUserDefinedProperty(propertyName);
        }
    }

    /// <summary>
    /// Centralized method for removing column from DataGrid
    /// Can accept either InventorPropertyName or ColumnHeader (for backward compatibility with Adorner drag)
    /// </summary>
    public void RemoveDataGridColumn(string propertyIdentifier, bool removeDataOnly = false)
    {
        // Try to find column by SortMemberPath first (stable identifier)
        var columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.SortMemberPath == propertyIdentifier);

        // If not found, try by Header (for Adorner drag compatibility)
        if (columnToRemove == null)
        {
            columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == propertyIdentifier);
        }

        if (columnToRemove != null)
        {
            var inventorPropertyName = columnToRemove.SortMemberPath;
            PartsDataGrid.Columns.Remove(columnToRemove);

            // Check if this is a UDP property
            if (!string.IsNullOrEmpty(inventorPropertyName) && PropertyMetadataRegistry.IsUserDefinedProperty(inventorPropertyName))
            {
                if (removeDataOnly)
                {
                    // Only remove UDP data (for Adorner)
                    RemoveUserDefinedIPropertyData(inventorPropertyName);
                }
                else
                {
                    // Complete UDP removal (for SelectIPropertyWindow)
                    RemoveUserDefinedIPropertyColumn(inventorPropertyName);
                }
            }
        }

        // Update property states in selection window
        var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>()
            .FirstOrDefault();
        selectIPropertyWindow?.UpdatePropertyStates();
    }

    /// <summary>
    /// Centralized method for creating text DataGrid columns
    /// </summary>
    private DataGridTextColumn CreateTextColumn(string header, string bindingPath, bool isSortable = true)
    {
        var binding = new Binding(bindingPath);

        return new DataGridTextColumn
        {
            Header = header,
            Binding = binding,
            SortMemberPath = bindingPath,
            CanUserSort = isSortable,
            ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style,
            IsReadOnly = true
        };
    }

    private Style CreateCellTagStyle(string propertyPath)
    {
        var style = PartsDataGrid.FindResource("DataGridCellStyle") is Style baseStyle
            ? new Style(typeof(DataGridCell), baseStyle)
            : new Style(typeof(DataGridCell));
        style.Setters.Add(new Setter(FrameworkElement.TagProperty, propertyPath));
        return style;
    }

    public void AddUserDefinedIPropertyColumn(string propertyName)
    {
        // Find user property in registry by original name
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InventorPropertyName == propertyName);
        var columnHeader = userProperty?.ColumnHeader ?? $"(User) {propertyName}";
        var internalName = userProperty?.InternalName ?? $"UDP_{propertyName}";

        // Check if column with this header already exists
        if (PartsDataGrid.Columns.Any(c => c.Header as string == columnHeader))
        {
            CustomMessageBox.Show(_localizationManager.GetString("Warning_ColumnAlreadyExists", columnHeader), _localizationManager.GetString("MessageBox_Warning"), MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        // Initialize empty values for all existing rows to avoid KeyNotFoundException
        // Use propertyName as key (stable, not localized)
        foreach (var partData in _partsData)
        {
            partData.AddUserDefinedProperty(propertyName, "");
        }

        // Create column with universal template and pass path via Tag
        var template = FindResource("EditableWithFxTemplate") as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException("EditableWithFxTemplate not found", "EditableWithFxTemplate");
        var column = new DataGridTemplateColumn
        {
            Header = columnHeader,
            CellTemplate = template,
            SortMemberPath = internalName,
            IsReadOnly = false,
            ClipboardContentBinding = new Binding($"UserDefinedProperties[{propertyName}]"),
            CellStyle = CreateCellTagStyle($"UserDefinedProperties[{propertyName}]")
        };

        PartsDataGrid.Columns.Add(column);

        // Fill data for new column
        FillPropertyData(internalName);
    }

    private void AboutButton_Click(object sender, RoutedEventArgs e)
    {
        var aboutWindow = new AboutWindow();
        aboutWindow.ShowDialog();
    }

    private void UpdateNoColumnsOverlayVisibility()
    {
        if (NoColumnsOverlay != null)
        {
            NoColumnsOverlay.Visibility = PartsDataGrid.Columns.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void PartsDataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Delete && PartsDataGrid.SelectedItems.Count > 0)
        {
            RemoveSelectedRows_Click(this, new RoutedEventArgs());
            e.Handled = true;
        }
    }

    private void ResetQuantity_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0) return;

        var overriddenItems = selectedItems.Where(item => item.IsOverridden).ToList();
        if (overriddenItems.Count == 0)
        {
            CustomMessageBox.Show(_localizationManager.GetString("Info_NoOverriddenQuantities"), _localizationManager.GetString("MessageBox_Information"),
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var message = overriddenItems.Count == 1
            ? string.Format(_localizationManager.GetString("Message_ResetQuantitySingle"),
                overriddenItems[0].PartNumber, overriddenItems[0].OriginalQuantity)
            : string.Format(_localizationManager.GetString("Message_ResetQuantityMultiple"),
                overriddenItems.Count);

        var result = CustomMessageBox.Show(message, _localizationManager.GetString("MessageBox_Confirmation"), MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        foreach (var item in overriddenItems)
        {
            item.Quantity = item.OriginalQuantity;
        }
    }

    /// <summary>
    /// Initialize language selection in ComboBox
    /// </summary>
    private void InitializeLanguageSelection()
    {
        var currentCulture = _localizationManager.CurrentCulture;
        var selectedLanguage = SupportedLanguages.All
            .FirstOrDefault(lang => lang.Culture.Name == currentCulture.Name)
            ?? SupportedLanguages.Russian;

        LanguageComboBox.SelectedItem = selectedLanguage;
    }

    /// <summary>
    /// Language change handler
    /// </summary>
    private void LanguageComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (sender is System.Windows.Controls.ComboBox comboBox && comboBox.SelectedItem is LanguageInfo selectedLanguage)
        {
            _localizationManager.CurrentCulture = selectedLanguage.Culture;
        }
    }

    private void ThemeToggleButton_Changed(object sender, RoutedEventArgs e)
    {
        if (ThemeToggleButton is null) return;

        var isDarkTheme = ThemeToggleButton.IsChecked == true;
        var themeFileName = isDarkTheme ? "DarkTheme.xaml" : "ColorResources.xaml";
        var themeUri = new Uri($"Styles/{themeFileName}", UriKind.Relative);

        var mergedDictionaries = System.Windows.Application.Current.Resources.MergedDictionaries;

        var existingTheme = mergedDictionaries.FirstOrDefault(d =>
            d.Source?.OriginalString.Contains("ColorResources.xaml") == true ||
            d.Source?.OriginalString.Contains("DarkTheme.xaml") == true);

        if (existingTheme != null)
        {
            var index = mergedDictionaries.IndexOf(existingTheme);
            mergedDictionaries.RemoveAt(index);
            mergedDictionaries.Insert(index, new ResourceDictionary { Source = themeUri });
        }
    }

    /// <summary>
    /// Updates button texts when language changes
    /// </summary>
    private void UpdateButtonTexts()
    {
        // Update button texts
        ScanButton.Content = _localizationManager.GetString(_isScanning ? "UIState_Cancel" : "UIState_Scan");
        ExportButton.Content = _localizationManager.GetString(_isExporting ? "UIState_Cancel" : "UIState_Export");

        // Update status text only if it doesn't contain dynamic data
        var currentText = ProgressLabelRun.Text;
        if (string.IsNullOrEmpty(currentText) ||
            (!currentText.Contains("/") && !currentText.Contains("(") && !currentText.Contains("-")))
        {
            // Determine which status text should be displayed
            if (_isScanning && _operationCts?.IsCancellationRequested == true)
            {
                ProgressLabelRun.Text = _localizationManager.GetString("UIState_Cancelling");
            }
            else if (_isScanning)
            {
                ProgressLabelRun.Text = _localizationManager.GetString("UIState_PreparingScan");
            }
            else if (_isExporting && _operationCts?.IsCancellationRequested == true)
            {
                ProgressLabelRun.Text = _localizationManager.GetString("UIState_CancellingExport");
            }
            else if (_isExporting)
            {
                ProgressLabelRun.Text = _localizationManager.GetString("UIState_ExportingData");
            }
            else if (_partsData.Count == 0)
            {
                ProgressLabelRun.Text = _localizationManager.GetString("UIState_NoDocumentSelected");
            }
        }
    }

    /// <summary>
    /// Updates column headers in DataGrid when language changes
    /// </summary>
    private void UpdatePresetPropertiesLocalization()
    {
        foreach (var column in PartsDataGrid.Columns.Where(c => c.SortMemberPath is string { Length: > 0 }))
        {
            var internalName = column.SortMemberPath!;
            var propertyDef = PropertyMetadataRegistry.GetPropertyByInternalName(internalName);
            if (propertyDef != null)
            {
                column.Header = propertyDef.ColumnHeader;
            }
        }
    }
}
