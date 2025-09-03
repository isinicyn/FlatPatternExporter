using System.Collections.ObjectModel;

namespace FlatPatternExporter;

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
    }

    /// <summary>
    /// Тип свойства
    /// </summary>
    public enum PropertyType
    {
        IProperty,        // Стандартное свойство Inventor
        Document,         // Свойство документа (не iProperty)
        System           // Системное свойство приложения
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
            ColumnTemplate = "EditableQuantityTemplate"
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
            Type = PropertyType.Document
        },
        ["Thickness"] = new PropertyDefinition
        {
            InternalName = "Thickness",
            DisplayName = "Толщина",
            ColumnHeader = "Толщина",
            Category = "Документ",
            Type = PropertyType.Document,
            RequiresRounding = true,
            RoundingDecimals = 1
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
            IsEditable = true
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
            IsEditable = true
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
            IsEditable = true
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
            IsEditable = true
        },
        ["Material"] = new PropertyDefinition
        {
            InternalName = "Material",
            DisplayName = "Материал",
            ColumnHeader = "Материал",
            Category = "Материал и отделка",
            Type = PropertyType.IProperty,
            PropertySetName = "Design Tracking Properties",
            InventorPropertyName = "Material"
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
            IsEditable = true
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
            RoundingDecimals = 2
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
            RoundingDecimals = 2
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
            RoundingDecimals = 2
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
            RoundingDecimals = 2
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
        }
    };

}