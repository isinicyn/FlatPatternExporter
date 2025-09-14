using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using Inventor;

namespace FlatPatternExporter.Core;
/// <summary>
/// Класс для централизованного управления доступом к свойствам документов Inventor.
/// Обеспечивает единый интерфейс для работы с iProperty и обычными свойствами документов.
/// </summary>
public class PropertyManager
{
    private readonly Document _document;
    public static readonly string SheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";

    public PropertyManager(Document document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>
    /// Получает маппинг для Inventor Property
    /// </summary>
    private static (string SetName, string InventorName) GetInventorMapping(string internalName)
    {
        // Сначала проверяем стандартные свойства
        if (PropertyMetadataRegistry.Properties.TryGetValue(internalName, out var def))
        {
            if (def.Type == PropertyMetadataRegistry.PropertyType.IProperty)
            {
                return (def.PropertySetName!, def.InventorPropertyName!);
            }
        }
        
        // Проверяем пользовательские свойства по InternalName с префиксом
        var userProperty = PropertyMetadataRegistry.GetPropertyByInternalName(internalName);
        if (userProperty != null && userProperty.Type == PropertyMetadataRegistry.PropertyType.UserDefined)
        {
            return (userProperty.PropertySetName!, userProperty.InventorPropertyName!);
        }
        
        // Если это UDP_ префикс, но свойство не найдено в реестре, извлекаем оригинальное имя
        if (PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
        {
            var originalName = PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(internalName);
            return ("User Defined Properties", originalName);
        }
        
        // Если свойство не найдено - это ошибка в коде
        throw new ArgumentException($"Неизвестное свойство: {internalName}");
    }

    /// <summary>
    /// Получает объект свойства по внутреннему имени с использованием централизованной системы метаданных.
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
            Debug.WriteLine($"Ошибка доступа к свойству '{ourName}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Получает значение или выражение свойства по внутреннему имени.
    /// </summary>
    /// <param name="ourName">Внутреннее имя свойства.</param>
    /// <param name="getExpression">Если true, возвращает выражение; иначе - значение.</param>
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

            // Получаем метаданные свойства
            PropertyMetadataRegistry.PropertyDefinition? definition = null;
            if (PropertyMetadataRegistry.Properties.TryGetValue(ourName, out var def))
            {
                definition = def;
            }
            if (definition != null)
            {
                // Применяем округление для числовых значений через централизованный метод
                if (definition.RequiresRounding && double.TryParse(result, out var numericValue))
                {
                    result = PropertyMetadataRegistry.FormatValue(ourName, numericValue);
                }

                // Применяем сопоставления значений
                if (definition.ValueMappings != null && definition.ValueMappings.TryGetValue(result, out var mappedValue))
                {
                    result = mappedValue;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Проверяет, является ли свойство expression-ом по внутреннему имени.
    /// </summary>
    public bool IsMappedPropertyExpression(string ourName)
    {
        var prop = GetPropertyObject(ourName);
        if (prop == null) return false;

        return !string.IsNullOrEmpty(prop.Expression) && prop.Expression.StartsWith("=");
    }

    /// <summary>
    /// Устанавливает значение свойства по внутреннему имени.
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
            Debug.WriteLine($"Ошибка установки значения свойства '{ourName}': {ex.Message}");
            System.Windows.MessageBox.Show($"Не удалось обновить свойство '{ourName}'.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Устанавливает выражение свойства по внутреннему имени.
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
            Debug.WriteLine($"Ошибка установки выражения свойства '{ourName}': {ex.Message}");
            System.Windows.MessageBox.Show($"Не удалось обновить выражение свойства '{ourName}'.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // === Методы для работы с не-iProperty свойствами ===

    /// <summary>
    /// Получает имя файла без расширения
    /// </summary>
    public string GetFileName()
    {
        try
        {
            return System.IO.Path.GetFileNameWithoutExtension(_document.FullFileName);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка получения имени файла: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Получает полный путь к файлу
    /// </summary>
    public string GetFullFileName()
    {
        try
        {
            return _document.FullFileName;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка получения полного пути к файлу: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Получает имя состояния модели
    /// </summary>
    public string GetModelState()
    {
        try
        {
            return _document.ModelStateName ?? "";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка получения состояния модели: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Получает толщину листового металла (только для деталей из листового металла)
    /// </summary>
    public string GetThickness()
    {
        try
        {
            if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
            {
                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var thicknessParam = smCompDef.Thickness;
                var thicknessValue = (double)thicknessParam.Value * 10; // Переводим в мм
                
                // Используем централизованное форматирование
                return PropertyMetadataRegistry.FormatValue("Thickness", thicknessValue);
            }
            return "0.0";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка получения толщины: {ex.Message}");
            return "0.0";
        }
    }

    /// <summary>
    /// Проверяет, имеет ли деталь развертку (только для деталей из листового металла)
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
            Debug.WriteLine($"Ошибка проверки наличия развертки: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Проверяет, является ли состояние модели документа основным (kPrimaryModelStateType = 118017)
    /// </summary>
    public bool IsPrimaryModelState()
    {
        try
        {
            // Получаем активное состояние модели из ComponentDefinition
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
            
            return true; // Для других типов документов считаем основным состоянием
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка проверки типа состояния модели: {ex.Message}");
            return true; // По умолчанию считаем основным состоянием
        }
    }

}