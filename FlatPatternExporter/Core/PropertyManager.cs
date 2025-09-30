using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using FlatPatternExporter.Services;
using Inventor;

namespace FlatPatternExporter.Core;
/// <summary>
/// Class for centralized management of access to Inventor document properties.
/// Provides a unified interface for working with iProperty and regular document properties.
/// </summary>
public class PropertyManager(Document document)
{
    private readonly Document _document = document ?? throw new ArgumentNullException(nameof(document));
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

        try
        {
            var propSet = _document.PropertySets[setName];
            return propSet[inventorName];
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property access error '{ourName}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets value or expression of property by internal name.
    /// </summary>
    /// <param name="ourName">Internal property name.</param>
    /// <param name="getExpression">If true, returns expression; otherwise - value.</param>
    public string GetMappedProperty(string ourName, bool getExpression = false)
    {
        var prop = GetPropertyObject(ourName);
        if (prop == null) return "";

        string result;
        if (getExpression)
        {
            result = prop.Expression ?? "";
        }
        else
        {
            result = prop.Value?.ToString() ?? "";

            // Get property metadata
            PropertyMetadataRegistry.PropertyDefinition? definition = null;
            if (PropertyMetadataRegistry.Properties.TryGetValue(ourName, out var def))
            {
                definition = def;
            }
            if (definition != null)
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
        }

        return result;
    }

    /// <summary>
    /// Checks if property is an expression by internal name.
    /// </summary>
    public bool IsMappedPropertyExpression(string ourName)
    {
        var prop = GetPropertyObject(ourName);
        if (prop == null) return false;

        return !string.IsNullOrEmpty(prop.Expression) && prop.Expression.StartsWith('=');
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
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property value set error '{ourName}': {ex.Message}");
            System.Windows.MessageBox.Show(LocalizationManager.Instance.GetString("Error_PropertyUpdateFailed", ourName), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
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
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Property expression set error '{ourName}': {ex.Message}");
            System.Windows.MessageBox.Show(LocalizationManager.Instance.GetString("Error_ExpressionUpdateFailed", ourName), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // === Methods for working with non-iProperty properties ===

    /// <summary>
    /// Gets file name without extension
    /// </summary>
    public string GetFileName()
    {
        try
        {
            return System.IO.Path.GetFileNameWithoutExtension(_document.FullFileName);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"File name retrieval error: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Gets full file path
    /// </summary>
    public string GetFullFileName()
    {
        try
        {
            return _document.FullFileName;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Full file path retrieval error: {ex.Message}");
            return "";
        }
    }

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
            if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
            {
                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
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
            if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
            {
                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                return smCompDef.HasFlatPattern;
            }
            return false;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Flat pattern check error: {ex.Message}");
            return false;
        }
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
