using System.Diagnostics;
using System.Windows;
using Inventor;

namespace FlatPatternExporter
{
    /// <summary>
    /// Класс для централизованного управления доступом к свойствам документов Inventor.
    /// Обеспечивает единый интерфейс для работы с iProperty и обычными свойствами документов.
    /// </summary>
    public class PropertyManager
    {
        private readonly Document _document;
        public static readonly string SheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";

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

        public PropertyManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Получает значение свойства из указанного набора свойств
        /// </summary>
        /// <param name="setName">Имя набора свойств</param>
        /// <param name="propName">Имя свойства</param>
        /// <returns>Значение свойства в виде строки или пустая строка при ошибке</returns>
        public string GetProperty(string setName, string propName)
        {
            try
            {
                var propSet = _document.PropertySets[setName];
                var prop = propSet[propName];
                return prop.Value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка получения свойства '{propName}' из набора '{setName}': {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Проверяет, является ли значение свойства expression-ом
        /// </summary>
        /// <param name="setName">Имя набора свойств</param>
        /// <param name="propName">Имя свойства</param>
        /// <returns>true если свойство содержит expression</returns>
        public bool IsPropertyExpression(string setName, string propName)
        {
            try
            {
                var propSet = _document.PropertySets[setName];
                var prop = propSet[propName];
                
                // Expression в Inventor должен начинаться с символа "="
                return !string.IsNullOrEmpty(prop.Expression) && prop.Expression.StartsWith("=");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка проверки expression свойства '{propName}' из набора '{setName}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Устанавливает значение свойства в указанном наборе свойств
        /// </summary>
        /// <param name="setName">Имя набора свойств</param>
        /// <param name="propName">Имя свойства</param>
        /// <param name="value">Новое значение свойства</param>
        public void SetProperty(string setName, string propName, object value)
        {
            try
            {
                var propSet = _document.PropertySets[setName];
                var prop = propSet[propName];
                prop.Value = value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка установки свойства '{propName}' в наборе '{setName}': {ex.Message}");
                System.Windows.MessageBox.Show($"Не удалось обновить свойство '{propName}'.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Получает значение свойства по внутреннему имени с использованием маппинга
        /// </summary>
        /// <param name="ourName">Внутреннее имя свойства</param>
        /// <returns>Значение свойства в виде строки</returns>
        public string GetMappedProperty(string ourName)
        {
            if (PropertyMapping.TryGetValue(ourName, out var mapping))
            {
                return GetProperty(mapping.SetName, mapping.InventorName);
            }

            // Для пользовательских свойств проверяем только набор user-defined
            return GetProperty("Inventor User Defined Properties", ourName);
        }

        /// <summary>
        /// Проверяет, является ли свойство expression-ом по внутреннему имени с использованием маппинга
        /// </summary>
        /// <param name="ourName">Внутреннее имя свойства</param>
        /// <returns>true если свойство содержит expression</returns>
        public bool IsMappedPropertyExpression(string ourName)
        {
            if (PropertyMapping.TryGetValue(ourName, out var mapping))
            {
                return IsPropertyExpression(mapping.SetName, mapping.InventorName);
            }

            // Для пользовательских свойств проверяем только набор user-defined
            return IsPropertyExpression("Inventor User Defined Properties", ourName);
        }

        /// <summary>
        /// Устанавливает значение свойства по внутреннему имени с использованием маппинга
        /// </summary>
        /// <param name="ourName">Внутреннее имя свойства</param>
        /// <param name="value">Новое значение свойства</param>
        public void SetMappedProperty(string ourName, object value)
        {
            if (PropertyMapping.TryGetValue(ourName, out var mapping))
            {
                SetProperty(mapping.SetName, mapping.InventorName, value);
                return;
            }

            // Для пользовательских свойств
            SetProperty("Inventor User Defined Properties", ourName, value);
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
                return string.Empty;
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
                return string.Empty;
            }
        }

        /// <summary>
        /// Получает имя состояния модели
        /// </summary>
        public string GetModelState()
        {
            try
            {
                return _document.ModelStateName ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка получения состояния модели: {ex.Message}");
                return string.Empty;
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
    }
}