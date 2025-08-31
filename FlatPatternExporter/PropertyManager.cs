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

        /// <summary>
        /// Список редактируемых свойств, которые могут быть выражениями
        /// </summary>
    private static readonly HashSet<string> EditableProperties =
    [
            // Design Tracking Properties
            "Authority", "CatalogWebLink", "CheckedBy", "CostCenter", "Description", 
            "Designer", "DesignStatus", "Engineer", "EngApprovedBy", "MfgApprovedBy", 
            "PartNumber", "Project", "StockNumber", "UserStatus", "Vendor",
            
            // Summary Information  
            "Author", "Comments", "Keywords", "Revision", "Subject", "Title",
            
            // Document Summary Information
        "Category", "Company", "Manager"
    ];

        /// <summary>
        /// Маппинг внутренних имен свойств на соответствующие наборы и имена в Inventor
        /// </summary>
        private static readonly Dictionary<string, (string SetName, string InventorName)> PropertyMapping = new()
        {
            // Summary Information
            { "Author", ("Summary Information", "Author") },
            { "Revision", ("Summary Information", "Revision Number") },
            { "Title", ("Summary Information", "Title") },
            { "Subject", ("Summary Information", "Subject") },
            { "Keywords", ("Summary Information", "Keywords") },
            { "Comments", ("Summary Information", "Comments") },

            // Document Summary Information
            { "Category", ("Document Summary Information", "Category") },
            { "Manager", ("Document Summary Information", "Manager") },
            { "Company", ("Document Summary Information", "Company") },

            // Design Tracking Properties
            { "PartNumber", ("Design Tracking Properties", "Part Number") },
            { "Description", ("Design Tracking Properties", "Description") },
            { "Material", ("Design Tracking Properties", "Material") },
            { "Project", ("Design Tracking Properties", "Project") },
            { "StockNumber", ("Design Tracking Properties", "Stock Number") },
            { "CreationTime", ("Design Tracking Properties", "Creation Time") },
            { "CostCenter", ("Design Tracking Properties", "Cost Center") },
            { "CheckedBy", ("Design Tracking Properties", "Checked By") },
            { "EngApprovedBy", ("Design Tracking Properties", "Engr Approved By") },
            { "UserStatus", ("Design Tracking Properties", "User Status") },
            { "CatalogWebLink", ("Design Tracking Properties", "Catalog Web Link") },
            { "Vendor", ("Design Tracking Properties", "Vendor") },
            { "MfgApprovedBy", ("Design Tracking Properties", "Mfg Approved By") },
            { "DesignStatus", ("Design Tracking Properties", "Design Status") },
            { "Designer", ("Design Tracking Properties", "Designer") },
            { "Engineer", ("Design Tracking Properties", "Engineer") },
            { "Authority", ("Design Tracking Properties", "Authority") },
            { "Mass", ("Design Tracking Properties", "Mass") },
            { "SurfaceArea", ("Design Tracking Properties", "SurfaceArea") },
            { "Volume", ("Design Tracking Properties", "Volume") },
            { "SheetMetalRule", ("Design Tracking Properties", "Sheet Metal Rule") },
            { "FlatPatternWidth", ("Design Tracking Properties", "Flat Pattern Width") },
            { "FlatPatternLength", ("Design Tracking Properties", "Flat Pattern Length") },
            { "FlatPatternArea", ("Design Tracking Properties", "Flat Pattern Area") },
            { "Appearance", ("Design Tracking Properties", "Appearance") }
        };

        /// <summary>
        /// Словарь сопоставлений значений свойств для преобразования внутренних значений в читаемые
        /// </summary>
    private static readonly Dictionary<string, Dictionary<string, string>> ValueMappings = new()
    {
        ["DesignStatus"] = new()
        {
            ["1"] = "Разработка",
            ["2"] = "Утверждение",
            ["3"] = "Завершен"
        }
    };

        /// <summary>
        /// Свойства, которые требуют округления до двух знаков после запятой
        /// </summary>
    private static readonly HashSet<string> NumericPropertiesForRounding =
    [
            "FlatPatternLength",
            "FlatPatternWidth", 
            "FlatPatternArea",
            "Mass",
            "Volume",
        "SurfaceArea"
    ];

        public PropertyManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Получает объект свойства по внутреннему имени с использованием маппинга или из пользовательских свойств.
        /// </summary>
        private Inventor.Property? GetPropertyObject(string ourName)
        {
            string setName;
            string inventorName;

        if (PropertyMapping.TryGetValue(ourName, out var mapping))
        {
            setName = mapping.SetName;
            inventorName = mapping.InventorName;
        }
        else
        {
            setName = "Inventor User Defined Properties";
            inventorName = ourName;
        }

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

            // Округляем числовые значения
            if (NumericPropertiesForRounding.Contains(ourName) && double.TryParse(result, out var numericValue))
            {
                result = Math.Round(numericValue, 2).ToString("F2");
            }

            // Применяем сопоставления значений
            if (ValueMappings.TryGetValue(ourName, out var mappings) && mappings.TryGetValue(result, out var mappedValue))
            {
                result = mappedValue;
            }
            }

            return result;
        }

        /// <summary>
        /// Получает выражение свойства по внутреннему имени с использованием маппинга
        /// </summary>
        public string GetMappedPropertyExpression(string ourName)
        {
            return GetMappedProperty(ourName, getExpression: true);
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
                    return Math.Round((double)thicknessParam.Value * 10, 1); // Переводим в мм и округляем до одной десятой
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
        /// Получает список предустановленных свойств iProperty на основе PropertyMapping
        /// </summary>
        public static ObservableCollection<PresetIProperty> GetPresetProperties()
        {
        ObservableCollection<PresetIProperty> presetProperties =
        [
                // Системные свойства приложения
                new() { ColumnHeader = "Обр.", ListDisplayName = "Статус обработки", InventorPropertyName = "ProcessingStatus", Category = "Системные" },
                new() { ColumnHeader = "ID", ListDisplayName = "Нумерация", InventorPropertyName = "Item", Category = "Системные" },

                // Свойства документа (не iProperty)
                new() { ColumnHeader = "Имя файла", ListDisplayName = "Имя файла", InventorPropertyName = "FileName", Category = "Документ" },
                new() { ColumnHeader = "Полное имя файла", ListDisplayName = "Полное имя файла", InventorPropertyName = "FullFileName", Category = "Документ" },
                new() { ColumnHeader = "Состояние модели", ListDisplayName = "Состояние модели", InventorPropertyName = "ModelState", Category = "Документ" },
                new() { ColumnHeader = "Толщина", ListDisplayName = "Толщина", InventorPropertyName = "Thickness", Category = "Документ" },
                new() { ColumnHeader = "Изобр. детали", ListDisplayName = "Изображение детали", InventorPropertyName = "Preview", Category = "Документ" },

                // Количество и обработка
                new() { ColumnHeader = "Кол.", ListDisplayName = "Количество", InventorPropertyName = "Quantity", Category = "Количество" },
            new() { ColumnHeader = "Изобр. развертки", ListDisplayName = "Изображение развертки", InventorPropertyName = "DxfPreview", Category = "Обработка" }
        ];

            // Автоматическое добавление свойств из PropertyMapping
            foreach (var mapping in PropertyMapping)
            {
            var category = mapping.Value.SetName switch
                {
                    "Summary Information" => "Summary Information",
                    "Document Summary Information" => "Document Summary Information", 
                    "Design Tracking Properties" => "Design Tracking Properties",
                    _ => "Прочие"
            };

            var columnHeader = GetDisplayNameForProperty(mapping.Key);
            var listDisplayName = GetDisplayNameForProperty(mapping.Key);
                
            presetProperties.Add(new() 
                { 
                    ColumnHeader = columnHeader,        // Заголовок колонки в DataGrid
                    ListDisplayName = listDisplayName, // Отображение в списке выбора
                    InventorPropertyName = mapping.Key, // Ключ для PropertyMapping
                    Category = category 
            });
            }

            return presetProperties;
        }

    /// <summary>
    /// Проверяет, является ли свойство редактируемым (может быть выражением)
    /// </summary>
    public static bool IsEditableProperty(string propertyName)
    {
        return EditableProperties.Contains(propertyName);
    }

    /// <summary>
    /// Возвращает коллекцию всех редактируемых свойств
    /// </summary>
    public static IEnumerable<string> GetEditableProperties()
    {
        return EditableProperties;
    }

        /// <summary>
        /// Получает русское отображаемое имя для свойства
        /// </summary>
        private static string GetDisplayNameForProperty(string propertyKey)
        {
            return propertyKey switch
            {
                "PartNumber" => "Обозначение",
                "Description" => "Наименование", 
                "Material" => "Материал",
                "Author" => "Автор",
                "Revision" => "Ревизия",
                "Title" => "Название",
                "Subject" => "Тема",
                "Keywords" => "Ключевые слова",
                "Comments" => "Примечание",
                "Category" => "Категория",
                "Manager" => "Менеджер",
                "Company" => "Компания",
                "Project" => "Проект",
                "StockNumber" => "Инвентарный номер",
                "CreationTime" => "Время создания",
                "CostCenter" => "Сметчик",
                "CheckedBy" => "Проверил",
                "EngApprovedBy" => "Нормоконтроль",
                "UserStatus" => "Статус",
                "CatalogWebLink" => "Веб-ссылка",
                "Vendor" => "Поставщик",
                "MfgApprovedBy" => "Утвердил",
                "DesignStatus" => "Статус разработки",
                "Designer" => "Проектировщик",
                "Engineer" => "Инженер",
                "Authority" => "Нач. отдела",
                "Mass" => "Масса",
                "SurfaceArea" => "Площадь поверхности",
                "Volume" => "Объем",
                "SheetMetalRule" => "Правило ЛМ",
                "FlatPatternWidth" => "Ширина развертки",
                "FlatPatternLength" => "Длина развертки",
                "FlatPatternArea" => "Площадь развертки",
                "Appearance" => "Отделка",
                _ => propertyKey
        };
    }
}