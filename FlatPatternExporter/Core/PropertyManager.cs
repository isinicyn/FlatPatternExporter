using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;
using Inventor;

namespace FlatPatternExporter.Core;
/// <summary>
/// Class for centralized management of access to Inventor document properties.
/// Provides a unified interface for working with iProperty and regular document properties.
/// </summary>
public class PropertyManager(Document document)
{
    private readonly Document _document = document ?? throw new ArgumentNullException(nameof(document));
    // Each lookup is a cross-process COM call, so resolved objects are reused for the lifetime of the manager
    private readonly Dictionary<string, PropertySet?> _propertySets = [];
    private PropertySets? _documentPropertySets;
    private SheetMetalComponentDefinition? _sheetMetalDefinition;
    // Values read by GetMappedProperties (null when the read failed); their names also tell missing properties apart
    private readonly Dictionary<string, Dictionary<string, object?>?> _propertySetValues = [];
    public static readonly string SheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";

    /// <summary>
    /// Gets the mapping for Inventor Property
    /// </summary>
    private static (string SetName, string InventorName) GetInventorMapping(string internalName)
    {
        // First check standard properties
        if (PropertyMetadataRegistry.Properties.TryGetValue(internalName, out var def))
        {
            if (def.Type == PropertyMetadataRegistry.PropertyType.IProperty)
            {
                return (def.PropertySetName!, def.InventorPropertyName!);
            }
        }

        // Check user-defined properties by InternalName with prefix
        var userProperty = PropertyMetadataRegistry.GetPropertyByInternalName(internalName);
        if (userProperty != null && userProperty.Type == PropertyMetadataRegistry.PropertyType.UserDefined)
        {
            return (userProperty.PropertySetName!, userProperty.InventorPropertyName!);
        }
        // If this is UDP_ prefix but property not found in registry, extract original name
        if (PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
        {
            var originalName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);
            return ("User Defined Properties", originalName);
        }
        // If property not found - this is a code error
        throw new ArgumentException($"Unknown property: {internalName}");
    }

    /// <summary>
    /// Gets property object by internal name using centralized metadata system.
    /// </summary>
    private Property? GetPropertyObject(string ourName)
    {
        var (setName, inventorName) = GetInventorMapping(ourName);

        // A property missing from an already read set is known without a failing COM call
        if (_propertySetValues.GetValueOrDefault(setName) is { } setValues && !setValues.ContainsKey(inventorName))
            return null;

        try
        {
            return GetPropertySet(setName)?[inventorName];
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property access error '{ourName}': {ex.Message}");
            return null;
        }
    }

    private PropertySet? GetPropertySet(string setName)
    {
        if (_propertySets.TryGetValue(setName, out var cached))
            return cached;

        PropertySet? propertySet = null;

        try
        {
            propertySet = (_documentPropertySets ??= _document.PropertySets)[setName];
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property set access error '{setName}': {ex.Message}");
        }

        return _propertySets[setName] = propertySet;
    }

    /// <summary>
    /// Gets value or expression of property by internal name.
    /// </summary>
    /// <param name="ourName">Internal property name.</param>
    /// <param name="getExpression">If true, returns expression; otherwise - value.</param>
    public string GetMappedProperty(string ourName, bool getExpression = false)
    {
        var prop = GetPropertyObject(ourName);

        if (getExpression)
            return prop?.Expression ?? "";

        return FormatMappedValue(ourName, prop?.Value);
    }

    /// <summary>
    /// Gets values of several properties by internal names with a single call per property set.
    /// </summary>
    public Dictionary<string, string> GetMappedProperties(IEnumerable<string> ourNames)
    {
        var result = new Dictionary<string, string>();

        foreach (var group in ourNames.Select(name => (Name: name, Mapping: GetInventorMapping(name))).GroupBy(x => x.Mapping.SetName))
        {
            if (!_propertySetValues.TryGetValue(group.Key, out var setValues))
                _propertySetValues[group.Key] = setValues = ReadPropertySetValues(group.Key);

            foreach (var (name, mapping) in group)
            {
                result[name] = setValues != null
                    ? FormatMappedValue(name, setValues.GetValueOrDefault(mapping.InventorName))
                    : GetMappedProperty(name);
            }
        }

        return result;
    }

    private Dictionary<string, object?>? ReadPropertySetValues(string setName)
    {
        if (GetPropertySet(setName) is not { } propertySet)
            return null;

        try
        {
            // The call goes through IDispatch, which requires non-null out arguments
            int[] ids = [];
            string[] names = [];
            object[] values = [];
            propertySet.GetPropertyInfo(out ids, out names, out values);

            // Property names are case-insensitive in Inventor; an empty set returns a single null entry
            var setValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < names.Length; i++)
            {
                if (names[i] != null)
                    setValues[names[i]] = values[i];
            }

            return setValues;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property set values read error '{setName}': {ex.Message}");
            return null;
        }
    }

    // A missing property is passed as null and ends up with the substitution value, like an empty one
    private static string FormatMappedValue(string ourName, object? value)
    {
        var result = value?.ToString() ?? "";

        if (PropertyMetadataRegistry.Properties.TryGetValue(ourName, out var definition))
        {
            // Apply rounding for numeric values through centralized method
            if (definition.RequiresRounding && double.TryParse(result, out var numericValue))
            {
                result = PropertyMetadataRegistry.FormatValue(ourName, numericValue);
            }

            // Apply value mappings
            if (definition.ValueMappings != null && definition.ValueMappings.TryGetValue(result, out var mappedValue))
            {
                result = mappedValue;
            }
        }

        // Apply substitution value if result is empty
        if (string.IsNullOrEmpty(result) && PropertyMetadataRegistry.PropertySubstitutions.TryGetValue(ourName, out var substitution))
        {
            result = substitution;
        }

        return result;
    }

    /// <summary>
    /// Checks if property is an expression by internal name.
    /// </summary>
    public bool IsMappedPropertyExpression(string ourName)
    {
        return GetPropertyObject(ourName)?.Expression?.StartsWith('=') == true;
    }

    /// <summary>
    /// Sets property value by internal name.
    /// </summary>
    public void SetMappedProperty(string ourName, object value)
    {
        var prop = GetPropertyObject(ourName);
        if (prop == null) return;

        try
        {
            prop.Value = value;
            _propertySetValues.Remove(GetInventorMapping(ourName).SetName);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property value set error '{ourName}': {ex.Message}");
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_PropertyUpdateFailed", ourName), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Sets property expression by internal name.
    /// </summary>
    public void SetMappedPropertyExpression(string ourName, string expression)
    {
        var prop = GetPropertyObject(ourName);
        if (prop == null) return;

        try
        {
            prop.Expression = expression;
            _propertySetValues.Remove(GetInventorMapping(ourName).SetName);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property expression set error '{ourName}': {ex.Message}");
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ExpressionUpdateFailed", ourName), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // === Methods for working with non-iProperty properties ===

    /// <summary>
    /// Gets model state name
    /// </summary>
    public string GetModelState()
    {
        try
        {
            return _document.ModelStateName ?? "";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Model state retrieval error: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Gets sheet metal thickness (for sheet metal parts only)
    /// </summary>
    public string GetThickness()
    {
        try
        {
            if (GetSheetMetalDefinition() is { } smCompDef)
            {
                var thicknessParam = smCompDef.Thickness;
                var thicknessValue = (double)thicknessParam.Value * 10; // Convert to mm
                // Use centralized formatting
                return PropertyMetadataRegistry.FormatValue("Thickness", thicknessValue);
            }
            return "0.0";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Thickness retrieval error: {ex.Message}");
            return "0.0";
        }
    }

    /// <summary>
    /// Checks if part has flat pattern (for sheet metal parts only)
    /// </summary>
    public bool HasFlatPattern()
    {
        try
        {
            return GetSheetMetalDefinition()?.HasFlatPattern ?? false;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Flat pattern check error: {ex.Message}");
            return false;
        }
    }

    private SheetMetalComponentDefinition? GetSheetMetalDefinition()
    {
        if (_sheetMetalDefinition == null && _document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
            _sheetMetalDefinition = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;

        return _sheetMetalDefinition;
    }

    /// <summary>
    /// Checks if document model state is primary (kPrimaryModelStateType = 118017)
    /// </summary>
    public bool IsPrimaryModelState()
    {
        try
        {
            // Get active model state from ComponentDefinition
            if (_document is PartDocument partDoc)
            {
                var activeModelState = partDoc.ComponentDefinition.ModelStates.ActiveModelState;
                return activeModelState.ModelStateType == ModelStateTypeEnum.kPrimaryModelStateType;
            }
            else if (_document is AssemblyDocument asmDoc)
            {
                var activeModelState = asmDoc.ComponentDefinition.ModelStates.ActiveModelState;
                return activeModelState.ModelStateType == ModelStateTypeEnum.kPrimaryModelStateType;
            }

            return true; // For other document types, consider it primary state
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Model state type check error: {ex.Message}");
            return true; // By default, consider it primary state
        }
    }

}
