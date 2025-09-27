using System.Collections.ObjectModel;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Services;

/// <summary>
/// Централизованная система метаданных свойств документов Inventor
/// </summary>
public static class PropertyMetadataRegistry
{
    /// <summary>
    /// Класс метаданных для одного свойства
    /// </summary>
    public class PropertyDefinition
    {
        public required string InternalName { get; init; }
        public required string DisplayName { get; init; }
        public string ColumnHeader { get; init; } = "";
        public string Category { get; init; } = "";
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
        
        // Вычисляемые свойства
        // TokenName для UDP содержит префикс для уникальности (UDP_Author != Author)
        public string TokenName => IsTokenizable ? InternalName : "";
        public string PlaceholderValue => IsTokenizable ? $"{{{DisplayName}}}" : "";
    }

    /// <summary>
    /// Тип свойства
    /// </summary>
    public enum PropertyType
    {
        IProperty,        // Стандартное свойство Inventor
        Document,         // Свойство документа (не iProperty)
        System,          // Системное свойство приложения
        UserDefined      // Пользовательское свойство iProperty
    }

    /// <summary>
    /// Основной реестр всех свойств
    /// </summary>
    public static readonly Dictionary<string, PropertyDefinition> Properties = new()
    {
        // ===== Системные свойства приложения =====
        ["ProcessingStatus"] = new PropertyDefinition
        {
            InternalName = "ProcessingStatus",
            DisplayName = "Статус обработки",
            ColumnHeader = "Обр.",
            Category = "Системные",
            Type = PropertyType.System,
            ColumnTemplate = "ProcessingStatusTemplate"
        },
        ["Item"] = new PropertyDefinition
        {
            InternalName = "Item",
            DisplayName = "Нумерация",
            ColumnHeader = "ID",
            Category = "Системные",
            Type = PropertyType.System,
            ColumnTemplate = "IDWithFlatPatternIndicatorTemplate",
            IsSearchable = false
        },
        ["Quantity"] = new PropertyDefinition
        {
            InternalName = "Quantity",
            DisplayName = "Количество",
            ColumnHeader = "Кол.",
            Category = "Количество",
            Type = PropertyType.System,
            ColumnTemplate = "EditableQuantityTemplate",
            IsTokenizable = true
        },
        
        // ===== Свойства документа (не iProperty) =====
        ["FileName"] = new PropertyDefinition
        {
            InternalName = "FileName",
            DisplayName = "Имя файла",
            ColumnHeader = "Имя файла",
            Category = "Документ",
            Type = PropertyType.Document
        },
        ["FullFileName"] = new PropertyDefinition
        {
            InternalName = "FullFileName",
            DisplayName = "Полное имя файла",
            ColumnHeader = "Полное имя файла",
            Category = "Документ",
            Type = PropertyType.Document
        },
        ["ModelState"] = new PropertyDefinition
        {
            InternalName = "ModelState",
            DisplayName = "Состояние модели",
            ColumnHeader = "Состояние модели",
            Category = "Документ",
            Type = PropertyType.Document,
            IsTokenizable = true
        },
        ["Thickness"] = new PropertyDefinition
        {
            InternalName = "Thickness",
            DisplayName = "Толщина",
            ColumnHeader = "Толщина",
            Category = "Документ",
            Type = PropertyType.Document,
            RequiresRounding = true,
            RoundingDecimals = 1,
            IsTokenizable = true
        },
        ["HasFlatPattern"] = new PropertyDefinition
        {
            InternalName = "HasFlatPattern",
            DisplayName = "Наличие развертки",
            ColumnHeader = "Развертка",
            Category = "Документ",
            Type = PropertyType.Document,
            IsSortable = true,
            IsSearchable = false
        },
        ["Preview"] = new PropertyDefinition
        {
            InternalName = "Preview",
            DisplayName = "Изображение детали",
            ColumnHeader = "Изобр. детали",
            Category = "Документ",
            Type = PropertyType.Document,
            ColumnTemplate = "PartImageTemplate",
            IsSortable = false,
            IsSearchable = false
        },
        ["DxfPreview"] = new PropertyDefinition
        {
            InternalName = "DxfPreview",
            DisplayName = "Изображение развертки",
            ColumnHeader = "Изобр. развертки",
            Category = "Обработка",
            Type = PropertyType.Document,
            ColumnTemplate = "DxfImageTemplate",
            IsSortable = false,
            IsSearchable = false
        },

        // ===== Summary Information =====
        ["Author"] = new PropertyDefinition
        {
            InternalName = "Author",
            DisplayName = "Автор",
            ColumnHeader = "Автор",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Author",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Revision"] = new PropertyDefinition
        {
            InternalName = "Revision",
            DisplayName = "Ревизия",
            ColumnHeader = "Ревизия",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Revision Number",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Title"] = new PropertyDefinition
        {
            InternalName = "Title",
            DisplayName = "Название",
            ColumnHeader = "Название",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Title",
            IsEditable = true
        },
        ["Subject"] = new PropertyDefinition
        {
            InternalName = "Subject",
            DisplayName = "Тема",
            ColumnHeader = "Тема",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Subject",
            IsEditable = true
        },
        ["Keywords"] = new PropertyDefinition
        {
            InternalName = "Keywords",
            DisplayName = "Ключевые слова",
            ColumnHeader = "Ключевые слова",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Keywords",
            IsEditable = true
        },
        ["Comments"] = new PropertyDefinition
        {
            InternalName = "Comments",
            DisplayName = "Примечание",
            ColumnHeader = "Примечание",
            Category = "Автор и редактирование",
            Type = PropertyType.IProperty,
            PropertySetName = "Summary Information",
            InventorPropertyName = "Comments",
            IsEditable = true
        },

        // ===== Document Summary Information =====
        ["Category"] = new PropertyDefinition
        {
            InternalName = "Category",
            DisplayName = "Категория",
            ColumnHeader = "Категория",
            Category = "Организационная информация",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Category",
            IsEditable = true
        },
        ["Manager"] = new PropertyDefinition
        {
            InternalName = "Manager",
            DisplayName = "Менеджер",
            ColumnHeader = "Менеджер",
            Category = "Организационная информация",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Manager",
            IsEditable = true
        },
        ["Company"] = new PropertyDefinition
        {
            InternalName = "Company",
            DisplayName = "Компания",
            ColumnHeader = "Компания",
            Category = "Организационная информация",
            Type = PropertyType.IProperty,
            PropertySetName = "Document Summary Information",
            InventorPropertyName = "Company",
            IsEditable = true
        },

        // ===== Design Tracking Properties =====
        ["PartNumber"] = new PropertyDefinition
        {
            InternalName = "PartNumber",
            DisplayName = "Обозначение",
            ColumnHeader = "Обозначение",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Part Number",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Description"] = new PropertyDefinition
        {
            InternalName = "Description",
            DisplayName = "Наименование",
            ColumnHeader = "Наименование",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Description",
            IsEditable = true,
            IsTokenizable = true
        },
        ["Material"] = new PropertyDefinition
        {
            InternalName = "Material",
            DisplayName = "Материал",
            ColumnHeader = "Материал",
            Category = "Материал и отделка",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Material",
            IsTokenizable = true
        },
        ["Project"] = new PropertyDefinition
        {
            InternalName = "Project",
            DisplayName = "Проект",
            ColumnHeader = "Проект",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Project",
            IsEditable = true,
            IsTokenizable = true
        },
        ["StockNumber"] = new PropertyDefinition
        {
            InternalName = "StockNumber",
            DisplayName = "Инвентарный номер",
            ColumnHeader = "Инвентарный номер",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Stock Number",
            IsEditable = true
        },
        ["CreationTime"] = new PropertyDefinition
        {
            InternalName = "CreationTime",
            DisplayName = "Время создания",
            ColumnHeader = "Время создания",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Creation Time"
        },
        ["CostCenter"] = new PropertyDefinition
        {
            InternalName = "CostCenter",
            DisplayName = "Сметчик",
            ColumnHeader = "Сметчик",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Cost Center",
            IsEditable = true
        },
        ["CheckedBy"] = new PropertyDefinition
        {
            InternalName = "CheckedBy",
            DisplayName = "Проверил",
            ColumnHeader = "Проверил",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Checked By",
            IsEditable = true
        },
        ["EngApprovedBy"] = new PropertyDefinition
        {
            InternalName = "EngApprovedBy",
            DisplayName = "Нормоконтроль",
            ColumnHeader = "Нормоконтроль",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Engr Approved By",
            IsEditable = true
        },
        ["UserStatus"] = new PropertyDefinition
        {
            InternalName = "UserStatus",
            DisplayName = "Статус",
            ColumnHeader = "Статус",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "User Status",
            IsEditable = true
        },
        ["CatalogWebLink"] = new PropertyDefinition
        {
            InternalName = "CatalogWebLink",
            DisplayName = "Веб-ссылка",
            ColumnHeader = "Веб-ссылка",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Catalog Web Link",
            IsEditable = true
        },
        ["Vendor"] = new PropertyDefinition
        {
            InternalName = "Vendor",
            DisplayName = "Поставщик",
            ColumnHeader = "Поставщик",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Vendor",
            IsEditable = true
        },
        ["MfgApprovedBy"] = new PropertyDefinition
        {
            InternalName = "MfgApprovedBy",
            DisplayName = "Утвердил",
            ColumnHeader = "Утвердил",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Mfg Approved By",
            IsEditable = true
        },
        ["DesignStatus"] = new PropertyDefinition
        {
            InternalName = "DesignStatus",
            DisplayName = "Статус разработки",
            ColumnHeader = "Статус разработки",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Design Status",
            IsEditable = true,
            ValueMappings = new Dictionary<string, string>
            {
                ["1"] = "Разработка",
                ["2"] = "Утверждение",
                ["3"] = "Завершен"
            }
        },
        ["Designer"] = new PropertyDefinition
        {
            InternalName = "Designer",
            DisplayName = "Проектировщик",
            ColumnHeader = "Проектировщик",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Designer",
            IsEditable = true
        },
        ["Engineer"] = new PropertyDefinition
        {
            InternalName = "Engineer",
            DisplayName = "Инженер",
            ColumnHeader = "Инженер",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Engineer",
            IsEditable = true
        },
        ["Authority"] = new PropertyDefinition
        {
            InternalName = "Authority",
            DisplayName = "Нач. отдела",
            ColumnHeader = "Нач. отдела",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Authority",
            IsEditable = true
        },
        ["Mass"] = new PropertyDefinition
        {
            InternalName = "Mass",
            DisplayName = "Масса",
            ColumnHeader = "Масса",
            Category = "Физические свойства",
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
            DisplayName = "Площадь поверхности",
            ColumnHeader = "Площадь поверхности",
            Category = "Листовой металл",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "SurfaceArea",
            RequiresRounding = true,
            RoundingDecimals = 2
        },
        ["Volume"] = new PropertyDefinition
        {
            InternalName = "Volume",
            DisplayName = "Объем",
            ColumnHeader = "Объем",
            Category = "Физические свойства",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Volume",
            RequiresRounding = true,
            RoundingDecimals = 2
        },
        ["SheetMetalRule"] = new PropertyDefinition
        {
            InternalName = "SheetMetalRule",
            DisplayName = "Правило ЛМ",
            ColumnHeader = "Правило ЛМ",
            Category = "Листовой металл",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Sheet Metal Rule"
        },
        ["FlatPatternWidth"] = new PropertyDefinition
        {
            InternalName = "FlatPatternWidth",
            DisplayName = "Ширина развертки",
            ColumnHeader = "Ширина развертки",
            Category = "Листовой металл",
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
            DisplayName = "Длина развертки",
            ColumnHeader = "Длина развертки",
            Category = "Листовой металл",
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
            DisplayName = "Площадь развертки",
            ColumnHeader = "Площадь развертки",
            Category = "Листовой металл",
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
            DisplayName = "Отделка",
            ColumnHeader = "Отделка",
            Category = "Материал и отделка",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Appearance"
        },
        ["Density"] = new PropertyDefinition
        {
            InternalName = "Density",
            DisplayName = "Плотность",
            ColumnHeader = "Плотность",
            Category = "Физические свойства",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Density",
            RequiresRounding = true,
            RoundingDecimals = 4
        },
        ["LastUpdatedWith"] = new PropertyDefinition
        {
            InternalName = "LastUpdatedWith",
            DisplayName = "Версия Inventor",
            ColumnHeader = "Версия Inventor",
            Category = "Техническая документация",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Last Updated With"
        }
    };

    /// <summary>
    /// Коллекция пользовательских свойств
    /// </summary>
    public static readonly ObservableCollection<PropertyDefinition> UserDefinedProperties = [];

    /// <summary>
    /// Создает определение пользовательского свойства
    /// </summary>
    public static PropertyDefinition CreateUserDefinedPropertyDefinition(string propertyName)
    {
        return new PropertyDefinition
        {
            InternalName = $"UDP_{propertyName}",
            DisplayName = propertyName,
            ColumnHeader = $"(Пользов.) {propertyName}",
            Category = "Пользовательские свойства",
            Type = PropertyType.UserDefined,
            PropertySetName = "User Defined Properties",
            InventorPropertyName = propertyName,
            IsEditable = true,
            IsTokenizable = true
        };
    }

    /// <summary>
    /// Добавляет пользовательское свойство в реестр
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
    /// Удаляет пользовательское свойство из реестра по InternalName
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
    /// Получает все свойства, которые могут использоваться как токены
    /// </summary>
    public static IEnumerable<PropertyDefinition> GetTokenizableProperties()
    {
        return Properties.Values.Where(p => p.IsTokenizable)
            .Concat(UserDefinedProperties.Where(p => p.IsTokenizable));
    }

    /// <summary>
    /// Получает список всех токенизируемых свойств
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
    /// Централизованное форматирование числовых значений согласно метаданным
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
    /// Получает InternalName по ColumnHeader из всех доступных свойств
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
    /// Получает PropertyDefinition по InternalName из всех доступных свойств
    /// </summary>
    public static PropertyDefinition? GetPropertyByInternalName(string internalName)
    {
        if (Properties.TryGetValue(internalName, out var presetProperty))
            return presetProperty;

        return UserDefinedProperties.FirstOrDefault(p => p.InternalName == internalName);
    }

    /// <summary>
    /// Проверяет, является ли InternalName пользовательским свойством
    /// </summary>
    public static bool IsUserDefinedProperty(string internalName)
    {
        return internalName.StartsWith("UDP_");
    }

    /// <summary>
    /// Извлекает Inventor имя из InternalName пользовательского свойства
    /// </summary>
    public static string GetInventorNameFromUserDefinedInternalName(string internalName)
    {
        return IsUserDefinedProperty(internalName) ? internalName[4..] : internalName;
    }

    /// <summary>
    /// Получает список предустановленных свойств iProperty из централизованной системы метаданных
    /// </summary>
    public static ObservableCollection<PresetIProperty> GetPresetProperties()
    {
        var presetProperties = new ObservableCollection<PresetIProperty>();

        // Добавляем стандартные свойства
        foreach (var prop in Properties.Values.OrderBy(p => p.Category).ThenBy(p => p.DisplayName))
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
        foreach (var userProp in UserDefinedProperties.OrderBy(p => p.DisplayName))
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
        return Properties.Values
            .Where(p => p.IsEditable)
            .Select(p => p.InternalName);
    }

}
