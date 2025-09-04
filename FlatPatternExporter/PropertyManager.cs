using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using Inventor;

namespace FlatPatternExporter;
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
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
        if (userProperty != null)
        {
            return (userProperty.PropertySetName!, userProperty.InventorPropertyName!);
        }
        
        // Если это UDP_ префикс, но свойство не найдено в реестре, извлекаем оригинальное имя
        if (internalName.StartsWith("UDP_"))
        {
            var originalName = internalName.Substring(4); // Убираем "UDP_"
            return ("User Defined Properties", originalName);
        }
        
        // Если свойство не найдено - это ошибка в коде
        throw new ArgumentException($"Неизвестное свойство: {internalName}");
    }

    /// <summary>
    /// Получает объект свойства по внутреннему имени с использованием централизованной системы метаданных.
    /// </summary>
    private Inventor.Property? GetPropertyObject(string ourName)
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
                // Округляем числовые значения
                if (definition.RequiresRounding && double.TryParse(result, out var numericValue))
                {
                    result = Math.Round(numericValue, definition.RoundingDecimals).ToString($"F{definition.RoundingDecimals}");
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
    public double GetThickness()
    {
        try
        {
            if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
            {
                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var thicknessParam = smCompDef.Thickness;
                PropertyMetadataRegistry.PropertyDefinition? definition = null;
                if (PropertyMetadataRegistry.Properties.TryGetValue("Thickness", out var def))
                {
                    definition = def;
                }
                var decimals = definition?.RoundingDecimals ?? 1;
                return Math.Round((double)thicknessParam.Value * 10, decimals); // Переводим в мм и округляем
            }
            return 0.0;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка получения толщины: {ex.Message}");
            return 0.0;
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

    /// <summary>
    /// Получает список предустановленных свойств iProperty из централизованной системы метаданных
    /// </summary>
    public static ObservableCollection<PresetIProperty> GetPresetProperties()
    {
        var presetProperties = new ObservableCollection<PresetIProperty>();

        // Добавляем стандартные свойства
        foreach (var prop in PropertyMetadataRegistry.Properties.Values.OrderBy(p => p.Category).ThenBy(p => p.DisplayName))
        {
            presetProperties.Add(new PresetIProperty
            {
                ColumnHeader = prop.ColumnHeader.Length > 0 ? prop.ColumnHeader : prop.DisplayName,
                ListDisplayName = prop.DisplayName,
                InventorPropertyName = prop.InternalName,
                Category = prop.Category
            });
        }

        // Добавляем пользовательские свойства
        foreach (var userProp in PropertyMetadataRegistry.UserDefinedProperties.OrderBy(p => p.DisplayName))
        {
            presetProperties.Add(new PresetIProperty
            {
                ColumnHeader = userProp.ColumnHeader.Length > 0 ? userProp.ColumnHeader : userProp.DisplayName,
                ListDisplayName = userProp.DisplayName,
                InventorPropertyName = userProp.InternalName,
                Category = userProp.Category
            });
        }

        return presetProperties;
    }


    /// <summary>
    /// Возвращает коллекцию всех редактируемых свойств из реестра
    /// </summary>
    public static IEnumerable<string> GetEditableProperties()
    {
        // Возвращаем только известные редактируемые свойства из реестра
        // Пользовательские свойства не включаются, так как они добавляются динамически
        return PropertyMetadataRegistry.Properties.Values
            .Where(p => p.IsEditable)
            .Select(p => p.InternalName);
    }


}