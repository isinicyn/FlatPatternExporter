using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Media.Imaging;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Models;

public class PartData : INotifyPropertyChanged
{
    private int item;
    private int quantity;
    private bool isOverridden;
    private bool isMultiplied;
    private BitmapImage? dxfPreview;
    private Enums.ProcessingStatus processingStatusEnum = Enums.ProcessingStatus.NotProcessed;
    private Dictionary<string, string> userDefinedProperties = [];
    private readonly Dictionary<string, bool> _isExpressionFlags = [];
    private int _expressionStateVersion;
    private bool _suppressExpressionVersion;
    private bool _expressionStateChangedWhileSuppressed;

    public PartData()
    {
        LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
    }

    private void OnLanguageChanged(object? sender, EventArgs e)
    {
        OnPropertyChanged(nameof(ProcessingStatus));
    }

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

    public int OriginalQuantity { get; set; }

    public bool IsOverridden
    {
        get => isOverridden;
        set
        {
            isOverridden = value;
            OnPropertyChanged();
        }
    }

    public bool IsMultiplied
    {
        get => isMultiplied;
        set
        {
            isMultiplied = value;
            OnPropertyChanged();
        }
    }

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

    public Enums.ProcessingStatus ProcessingStatusEnum
    {
        get => processingStatusEnum;
        set
        {
            if (processingStatusEnum != value)
            {
                processingStatusEnum = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProcessingStatus));
            }
        }
    }

    public string ProcessingStatus => ProcessingStatusEnum switch
    {
        Enums.ProcessingStatus.NotProcessed => LocalizationManager.Instance.GetString("ProcessingStatus_NotProcessed"),
        Enums.ProcessingStatus.Pending => LocalizationManager.Instance.GetString("ProcessingStatus_Pending"),
        Enums.ProcessingStatus.Success => LocalizationManager.Instance.GetString("ProcessingStatus_Success"),
        Enums.ProcessingStatus.Skipped => LocalizationManager.Instance.GetString("ProcessingStatus_Skipped"),
        _ => string.Empty
    };

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
    /// Checks if the specified property is an expression
    /// </summary>
    public bool IsPropertyExpression(string propertyName) =>
        _isExpressionFlags.TryGetValue(propertyName, out var isExpression) && isExpression;

    /// <summary>
    /// Sets the expression state for a property
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

    public event PropertyChangedEventHandler? PropertyChanged;

    public void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    public void AddUserDefinedProperty(string propertyName, string propertyValue)
    {
        UserDefinedProperties[propertyName] = propertyValue;
        OnPropertyChanged(nameof(UserDefinedProperties));
    }

    public void RemoveUserDefinedProperty(string propertyName) =>
        userDefinedProperties.Remove(propertyName);
}