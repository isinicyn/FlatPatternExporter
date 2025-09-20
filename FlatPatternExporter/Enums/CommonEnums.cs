using netDxf.Header;

namespace FlatPatternExporter.Enums;

public enum ExportFolderType
{
    ChooseFolder = 0,      // Указать папку в процессе экспорта
    ComponentFolder = 1,   // Папка компонента
    PartFolder = 2,        // Рядом с деталью
    ProjectFolder = 3,     // Папка проекта
    FixedFolder = 4        // Фиксированная папка
}

public enum ProcessingMethod
{
    Traverse = 0,          // Перебор
    BOM = 1               // Спецификация
}

public enum SplineReplacementType
{
    Lines = 0,            // Линии
    Arcs = 1              // Дуги
}

public enum AcadVersionType
{
    V2018 = 0,           // 2018
    V2013 = 1,           // 2013
    V2010 = 2,           // 2010
    V2007 = 3,           // 2007
    V2004 = 4,           // 2004
    V2000 = 5,           // 2000 (по умолчанию)
    R12 = 6              // R12
}

public enum ProcessingStatus
{
    NotProcessed,   // Не обработан (прозрачный)
    Success,        // Успешно экспортирован (зеленый)
    Error          // Ошибка экспорта (красный)
}

public enum DocumentType
{
    Assembly,
    Part,
    Invalid
}

public enum OperationType
{
    Scan,
    Export
}

public static class AcadVersionMapping
{
    private static readonly Dictionary<AcadVersionType, (string Name, DxfVersion? Dxf)> Map = new()
    {
        { AcadVersionType.V2018, ("2018", DxfVersion.AutoCad2018) },
        { AcadVersionType.V2013, ("2013", DxfVersion.AutoCad2013) },
        { AcadVersionType.V2010, ("2010", DxfVersion.AutoCad2010) },
        { AcadVersionType.V2007, ("2007", DxfVersion.AutoCad2007) },
        { AcadVersionType.V2004, ("2004", DxfVersion.AutoCad2004) },
        { AcadVersionType.V2000, ("2000", DxfVersion.AutoCad2000) },
        { AcadVersionType.R12,   ("R12",  null) }
    };

    public static string GetDisplayName(AcadVersionType type) => Map[type].Name;

    public static string GetTranslatorCode(AcadVersionType type) => Map[type].Name;

    public static bool SupportsOptimization(AcadVersionType type) => Map[type].Dxf.HasValue;

    public static DxfVersion? GetDxfVersion(AcadVersionType type) => Map[type].Dxf;
}

public static class SplineReplacementMapping
{
    private static readonly Dictionary<SplineReplacementType, string> Map = new()
    {
        { SplineReplacementType.Lines, "Линии" },
        { SplineReplacementType.Arcs, "Дуги" }
    };

    public static string GetDisplayName(SplineReplacementType type) => Map[type];
}