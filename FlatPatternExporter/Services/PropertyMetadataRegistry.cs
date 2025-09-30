using System.Collections.ObjectModel;
using System.ComponentModel;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Services;

/// <summary>
/// Centralized metadata system for Inventor document properties
/// </summary>
public static class PropertyMetadataRegistry
{
    /// <summary>
    /// Metadata class for a single property
    /// </summary>
    public class PropertyDefinition : INotifyPropertyChanged
    {
        private string? _localizationKeyPrefix;
        private string? _displayName;
        private string? _columnHeader;
        private string? _category;
        private string _internalName = "";

        public PropertyDefinition()
        {
            LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
        }

        protected PropertyDefinition(string internalName)
        {
            _internalName = internalName;
            LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
        }

        private void OnLanguageChanged(object? sender, EventArgs e)
        {
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(ColumnHeader));
            OnPropertyChanged(nameof(Category));
            OnPropertyChanged(nameof(PlaceholderValue));
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public string InternalName
        {
            get => _internalName;
            init => _internalName = value;
        }

        public string? LocalizationKeyPrefix
        {
            get => _localizationKeyPrefix;
            init => _localizationKeyPrefix = value;
        }

        public virtual string DisplayName
        {
            get => !string.IsNullOrEmpty(LocalizationKeyPrefix)
                ? LocalizationManager.Instance.GetString($"{LocalizationKeyPrefix}_DisplayName")
                : _displayName ?? "";
            init => _displayName = value;
        }

        public virtual string ColumnHeader
        {
            get => !string.IsNullOrEmpty(LocalizationKeyPrefix)
                ? LocalizationManager.Instance.GetString($"{LocalizationKeyPrefix}_ColumnHeader")
                : _columnHeader ?? "";
            init => _columnHeader = value;
        }

        public virtual string Category
        {
            get => !string.IsNullOrEmpty(LocalizationKeyPrefix)
                ? LocalizationManager.Instance.GetString($"{LocalizationKeyPrefix}_Category")
                : _category ?? "";
            init => _category = value;
        }

        public PropertyType Type { get; init; }
        public string? PropertySetName { get; init; }
        public string? InventorPropertyName { get; init; }
        public bool IsEditable { get; init; }
        public bool RequiresRounding { get; init; }
        public int RoundingDecimals { get; init; } = 2;
        public Dictionary<string, string>? ValueMappings { get; init; }
        public string? ColumnTemplate { get; init; }
        public bool IsSortable { get; init; } = true;
        public bool IsSearchable { get; init; } = true;
        public bool IsTokenizable { get; init; } = false;

        public string TokenName => IsTokenizable ? InternalName : "";
        public string PlaceholderValue => IsTokenizable ? $"{{{DisplayName}}}" : "";
    }

    /// <summary>
    /// Property type
    /// </summary>
    public enum PropertyType
    {
        IProperty,        // Standard Inventor property
        Document,         // Document property (not iProperty)
        System,          // Application system property
        UserDefined      // User-defined iProperty
    }

    /// <summary>
    /// Main registry of all properties
    /// </summary>
    public static readonly Dictionary<string, PropertyDefinition> Properties = new()
    {
        // ===== Application system properties =====
        ["ProcessingStatus"] = new PropertyDefinition
        {
            InternalName = "ProcessingStatus",
            LocalizationKeyPrefix = "Property_ProcessingStatus",
            Type = PropertyType.System,
            ColumnTemplate = "ProcessingStatusTemplate"
        },
        ["Item"] = new PropertyDefinition
        {
            InternalName = "Item",
            LocalizationKeyPrefix = "Property_Item",
            Type = PropertyType.System,
            ColumnTemplate = "IDWithFlatPatternIndicatorTemplate",
            IsSearchable = false
        },
        ["Quantity"] = new PropertyDefinition
        {
            InternalName = "Quantity",
            LocalizationKeyPrefix = "Property_Quantity",
            Type = PropertyType.System,
            ColumnTemplate = "EditableQuantityTemplate",
            IsTokenizable = true
        },

        // ===== Document properties (not iProperty) =====
        ["FileName"] = new PropertyDefinition
        {
            InternalName = "FileName",
            LocalizationKeyPrefix = "Property_FileName",
            Type = PropertyType.Document
        },
        ["FullFileName"] = new PropertyDefinition
        {
            InternalName = "FullFileName",
            LocalizationKeyPrefix = "Property_FullFileName",
            Type = PropertyType.Document
        },
        ["ModelState"] = new PropertyDefinition
        {
            InternalName = "ModelState",
            LocalizationKeyPrefix = "Property_ModelState",
            Type = PropertyType.Document,
            IsTokenizable = true
        },
        ["Thickness"] = new PropertyDefinition
        {
            InternalName = "Thickness",
            LocalizationKeyPrefix = "Property_Thickness",
            Type = PropertyType.Document,
            RequiresRounding = true,
            RoundingDecimals = 1,
            IsTokenizable = true
        },
        ["HasFlatPattern"] = new PropertyDefinition
        {
            InternalName = "HasFlatPattern",
            LocalizationKeyPrefix = "Property_HasFlatPattern",
            Type = PropertyType.Document,
            IsSearchable = false
        },
        ["Preview"] = new PropertyDefinition
        {
            InternalName = "Preview",
            LocalizationKeyPrefix = "Property_Preview",
            Type = PropertyType.Document,
            ColumnTemplate = "PartImageTemplate",
            IsSortable = false,
            IsSearchable = false
        },
        ["DxfPreview"] = new PropertyDefinition
        {
            InternalName = "DxfPreview",
            LocalizationKeyPrefix = "Property_DxfPreview",
            Type = PropertyType.Document,
            ColumnTemplate = "DxfImageTemplate",
            IsSortable = false,
            IsSearchable = false
        },

        // ===== Summary Information =====
        ["Author"] = new PropertyDefinition
        {
            InternalName = "Author",
            LocalizationKeyPrefix = "Property_Author",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Author",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Revision"] = new PropertyDefinition
        {
            InternalName = "Revision",
            LocalizationKeyPrefix = "Property_Revision",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Revision Number",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Title"] = new PropertyDefinition
        {
            InternalName = "Title",
            LocalizationKeyPrefix = "Property_Title",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Title",
            IsEditable = true
        },
        ["Subject"] = new PropertyDefinition
        {
            InternalName = "Subject",
            LocalizationKeyPrefix = "Property_Subject",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Subject",
            IsEditable = true
        },
        ["Keywords"] = new PropertyDefinition
        {
            InternalName = "Keywords",
            LocalizationKeyPrefix = "Property_Keywords",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Keywords",
            IsEditable = true
        },
        ["Comments"] = new PropertyDefinition
        {
            InternalName = "Comments",
            LocalizationKeyPrefix = "Property_Comments",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Comments",
            IsEditable = true
        },

        // ===== Document Summary Information =====
        ["Category"] = new PropertyDefinition
        {
            InternalName = "Category",
            LocalizationKeyPrefix = "Property_Category",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Category",
            IsEditable = true
        },
        ["Manager"] = new PropertyDefinition
        {
            InternalName = "Manager",
            LocalizationKeyPrefix = "Property_Manager",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Manager",
            IsEditable = true
        },
        ["Company"] = new PropertyDefinition
        {
            InternalName = "Company",
            LocalizationKeyPrefix = "Property_Company",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Company",
            IsEditable = true
        },

        // ===== Design Tracking Properties =====
        ["PartNumber"] = new PropertyDefinition
        {
            InternalName = "PartNumber",
            LocalizationKeyPrefix = "Property_PartNumber",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Part Number",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Description"] = new PropertyDefinition
        {
            InternalName = "Description",
            LocalizationKeyPrefix = "Property_Description",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Description",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Material"] = new PropertyDefinition
        {
            InternalName = "Material",
            LocalizationKeyPrefix = "Property_Material",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Material",
            IsTokenizable = true
        },
        ["Project"] = new PropertyDefinition
        {
            InternalName = "Project",
            LocalizationKeyPrefix = "Property_Project",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Project",
            IsEditable = true,
            IsTokenizable = true
        },
        ["StockNumber"] = new PropertyDefinition
        {
            InternalName = "StockNumber",
            LocalizationKeyPrefix = "Property_StockNumber",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Stock Number",
            IsEditable = true
        },
        ["CreationTime"] = new PropertyDefinition
        {
            InternalName = "CreationTime",
            LocalizationKeyPrefix = "Property_CreationTime",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Creation Time"
        },
        ["CostCenter"] = new PropertyDefinition
        {
            InternalName = "CostCenter",
            LocalizationKeyPrefix = "Property_CostCenter",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Cost Center",
            IsEditable = true
        },
        ["CheckedBy"] = new PropertyDefinition
        {
            InternalName = "CheckedBy",
            LocalizationKeyPrefix = "Property_CheckedBy",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Checked By",
            IsEditable = true
        },
        ["EngApprovedBy"] = new PropertyDefinition
        {
            InternalName = "EngApprovedBy",
            LocalizationKeyPrefix = "Property_EngApprovedBy",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Engr Approved By",
            IsEditable = true
        },
        ["UserStatus"] = new PropertyDefinition
        {
            InternalName = "UserStatus",
            LocalizationKeyPrefix = "Property_UserStatus",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "User Status",
            IsEditable = true
        },
        ["CatalogWebLink"] = new PropertyDefinition
        {
            InternalName = "CatalogWebLink",
            LocalizationKeyPrefix = "Property_CatalogWebLink",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Catalog Web Link",
            IsEditable = true
        },
        ["Vendor"] = new PropertyDefinition
        {
            InternalName = "Vendor",
            LocalizationKeyPrefix = "Property_Vendor",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Vendor",
            IsEditable = true
        },
        ["MfgApprovedBy"] = new PropertyDefinition
        {
            InternalName = "MfgApprovedBy",
            LocalizationKeyPrefix = "Property_MfgApprovedBy",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Mfg Approved By",
            IsEditable = true
        },
        ["DesignStatus"] = new PropertyDefinition
        {
            InternalName = "DesignStatus",
            LocalizationKeyPrefix = "Property_DesignStatus",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Design Status",
            IsEditable = true,
            ValueMappings = new Dictionary<string, string>
            {
                ["1"] = LocalizationManager.Instance.GetString("DesignStatus_Development"),
                ["2"] = LocalizationManager.Instance.GetString("DesignStatus_Approval"),
                ["3"] = LocalizationManager.Instance.GetString("DesignStatus_Completed")
            }
        },
        ["Designer"] = new PropertyDefinition
        {
            InternalName = "Designer",
            LocalizationKeyPrefix = "Property_Designer",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Designer",
            IsEditable = true
        },
        ["Engineer"] = new PropertyDefinition
        {
            InternalName = "Engineer",
            LocalizationKeyPrefix = "Property_Engineer",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Engineer",
            IsEditable = true
        },
        ["Authority"] = new PropertyDefinition
        {
            InternalName = "Authority",
            LocalizationKeyPrefix = "Property_Authority",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Authority",
            IsEditable = true
        },
        ["Mass"] = new PropertyDefinition
        {
            InternalName = "Mass",
            LocalizationKeyPrefix = "Property_Mass",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Mass",
            RequiresRounding = true,
            RoundingDecimals = 2,
            IsTokenizable = true
        },
        ["SurfaceArea"] = new PropertyDefinition
        {
            InternalName = "SurfaceArea",
            LocalizationKeyPrefix = "Property_SurfaceArea",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "SurfaceArea",
            RequiresRounding = true,
            RoundingDecimals = 2
        },
        ["Volume"] = new PropertyDefinition
        {
            InternalName = "Volume",
            LocalizationKeyPrefix = "Property_Volume",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Volume",
            RequiresRounding = true,
            RoundingDecimals = 2
        },
        ["SheetMetalRule"] = new PropertyDefinition
        {
            InternalName = "SheetMetalRule",
            LocalizationKeyPrefix = "Property_SheetMetalRule",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Sheet Metal Rule"
        },
        ["FlatPatternWidth"] = new PropertyDefinition
        {
            InternalName = "FlatPatternWidth",
            LocalizationKeyPrefix = "Property_FlatPatternWidth",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Flat Pattern Width",
            RequiresRounding = true,
            RoundingDecimals = 2,
            IsTokenizable = true
        },
        ["FlatPatternLength"] = new PropertyDefinition
        {
            InternalName = "FlatPatternLength",
            LocalizationKeyPrefix = "Property_FlatPatternLength",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Flat Pattern Length",
            RequiresRounding = true,
            RoundingDecimals = 2,
            IsTokenizable = true
        },
        ["FlatPatternArea"] = new PropertyDefinition
        {
            InternalName = "FlatPatternArea",
            LocalizationKeyPrefix = "Property_FlatPatternArea",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Flat Pattern Area",
            RequiresRounding = true,
            RoundingDecimals = 2,
            IsTokenizable = true
        },
        ["Appearance"] = new PropertyDefinition
        {
            InternalName = "Appearance",
            LocalizationKeyPrefix = "Property_Appearance",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Appearance"
        },
        ["Density"] = new PropertyDefinition
        {
            InternalName = "Density",
            LocalizationKeyPrefix = "Property_Density",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Density",
            RequiresRounding = true,
            RoundingDecimals = 4
        },
        ["LastUpdatedWith"] = new PropertyDefinition
        {
            InternalName = "LastUpdatedWith",
            LocalizationKeyPrefix = "Property_LastUpdatedWith",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Last Updated With"
        }
    };

    /// <summary>
    /// Collection of user-defined properties
    /// </summary>
    public static readonly ObservableCollection<PropertyDefinition> UserDefinedProperties = [];

    /// <summary>
    /// Dictionary of property substitutions (values used when property is empty or missing)
    /// Key: InternalName, Value: SubstitutionValue
    /// </summary>
    public static readonly Dictionary<string, string> PropertySubstitutions = [];

    /// <summary>
    /// Creates a user-defined property definition
    /// </summary>
    public static PropertyDefinition CreateUserDefinedPropertyDefinition(string propertyName)
    {
        return new UserDefinedPropertyDefinition(propertyName);
    }

    /// <summary>
    /// Special PropertyDefinition for user-defined properties with dynamic localization
    /// </summary>
    private class UserDefinedPropertyDefinition : PropertyDefinition
    {
        private readonly string _propertyName;

        public UserDefinedPropertyDefinition(string propertyName) : base($"UDP_{propertyName}")
        {
            _propertyName = propertyName;
            Type = PropertyType.UserDefined;
            PropertySetName = "User Defined Properties";
            InventorPropertyName = propertyName;
            IsEditable = true;
            IsTokenizable = true;
        }

        public override string DisplayName => _propertyName;

        public override string ColumnHeader =>
            $"{LocalizationManager.Instance.GetString("Property_UserDefined_Prefix")} {_propertyName}";

        public override string Category =>
            LocalizationManager.Instance.GetString("Property_UserDefined_Category");
    }

    /// <summary>
    /// Adds a user-defined property to the registry
    /// </summary>
    public static void AddUserDefinedProperty(string propertyName)
    {
        var internalName = $"UDP_{propertyName}";
        if (UserDefinedProperties.Any(p => p.InternalName == internalName))
            return;

        var userProperty = CreateUserDefinedPropertyDefinition(propertyName);
        UserDefinedProperties.Add(userProperty);
    }

    /// <summary>
    /// Removes a user-defined property from the registry by InternalName
    /// </summary>
    public static void RemoveUserDefinedProperty(string internalName)
    {
        var property = UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
        if (property != null)
        {
            UserDefinedProperties.Remove(property);
        }
    }

    /// <summary>
    /// Gets all properties that can be used as tokens
    /// </summary>
    public static IEnumerable<PropertyDefinition> GetTokenizableProperties()
    {
        return Properties.Values.Where(p => p.IsTokenizable)
            .Concat(UserDefinedProperties.Where(p => p.IsTokenizable));
    }

    /// <summary>
    /// Gets a list of all tokenizable properties
    /// </summary>
    public static Dictionary<string, string> GetAvailableTokens()
    {
        var tokens = new Dictionary<string, string>();
        
        foreach (var prop in GetTokenizableProperties())
        {
            var tokenName = prop.TokenName;
            if (!string.IsNullOrEmpty(tokenName))
            {
                tokens[tokenName] = prop.DisplayName;
            }
        }
        
        return tokens;
    }

    /// <summary>
    /// Centralized formatting of numeric values according to metadata
    /// </summary>
    public static string FormatValue(string propertyName, object? value)
    {
        if (value == null) return "";

        if (Properties.TryGetValue(propertyName, out var propertyDef) && 
            propertyDef.RequiresRounding && 
            value is double dValue)
        {
            return Math.Round(dValue, propertyDef.RoundingDecimals).ToString($"F{propertyDef.RoundingDecimals}");
        }

        return value.ToString() ?? "";
    }

    /// <summary>
    /// Gets InternalName by ColumnHeader from all available properties
    /// </summary>
    public static string? GetInternalNameByColumnHeader(string columnHeader)
    {
        var presetProperty = Properties.Values.FirstOrDefault(p => p.ColumnHeader == columnHeader);
        if (presetProperty != null)
            return presetProperty.InternalName;

        var userProperty = UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeader);
        return userProperty?.InternalName;
    }

    /// <summary>
    /// Gets PropertyDefinition by InternalName from all available properties
    /// </summary>
    public static PropertyDefinition? GetPropertyByInternalName(string internalName)
    {
        if (Properties.TryGetValue(internalName, out var presetProperty))
            return presetProperty;

        return UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
    }

    /// <summary>
    /// Checks if InternalName is a user-defined property
    /// </summary>
    public static bool IsUserDefinedProperty(string internalName)
    {
        return internalName.StartsWith("UDP_");
    }

    /// <summary>
    /// Extracts Inventor name from InternalName of user-defined property
    /// </summary>
    public static string GetInventorNameFromUserDefinedInternalName(string internalName)
    {
        return IsUserDefinedProperty(internalName) ? internalName[4..] : internalName;
    }

    /// <summary>
    /// Gets a list of preset iProperties from the centralized metadata system
    /// </summary>
    public static ObservableCollection<PresetIProperty> GetPresetProperties()
    {
        var presetProperties = new ObservableCollection<PresetIProperty>();

        // Add standard properties
        foreach (var prop in Properties.Values.OrderBy(p => p.Category).ThenBy(p => p.DisplayName))
        {
            presetProperties.Add(new PresetIProperty
            {
                InventorPropertyName = prop.InternalName
            });
        }

        // Add user-defined properties
        foreach (var userProp in UserDefinedProperties.OrderBy(p => p.DisplayName))
        {
            presetProperties.Add(new PresetIProperty
            {
                InventorPropertyName = userProp.InternalName
            });
        }

        return presetProperties;
    }

    /// <summary>
    /// Returns a collection of all editable properties from the registry
    /// </summary>
    public static IEnumerable<string> GetEditableProperties()
    {
        // Return only known editable properties from the registry
        // User-defined properties are not included as they are added dynamically
        return Properties.Values
            .Where(p => p.IsEditable)
            .Select(p => p.InternalName);
    }

}
