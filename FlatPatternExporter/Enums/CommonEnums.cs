using netDxf.Header;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Enums;

public enum ExportFolderType
{
    ChooseFolder = 0,      // Choose folder during export
    ComponentFolder = 1,   // Component folder
    PartFolder = 2,        // Next to part
    ProjectFolder = 3,     // Project folder
    FixedFolder = 4        // Fixed folder
}

public enum ProcessingMethod
{
    Traverse = 0,          // Traverse
    BOM = 1               // Bill of Materials
}

public enum SplineReplacementType
{
    Lines = 0,            // Lines
    Arcs = 1              // Arcs
}

public enum CsvDelimiterType
{
    Comma = 0,           // Comma (,)
    Semicolon = 1,       // Semicolon (;)
    Tab = 2,             // Tab (\t)
    Pipe = 3             // Pipe (|)
}

public enum AcadVersionType
{
    V2018 = 0,           // 2018
    V2013 = 1,           // 2013
    V2010 = 2,           // 2010
    V2007 = 3,           // 2007
    V2004 = 4,           // 2004
    V2000 = 5,           // 2000 (default)
    R12 = 6              // R12
}

public enum ProcessingStatus
{
    NotProcessed,   // Not processed (transparent)
    Pending,        // Pending export (yellow/orange)
    Success,        // Successfully exported (green)
    Skipped,        // Skipped (no flat pattern or error) (red)
    Interrupted     // Export was interrupted (gray)
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
    public static string GetDisplayName(SplineReplacementType type) => type switch
    {
        SplineReplacementType.Lines => LocalizationManager.Instance.GetString("SplineReplacement_Lines"),
        SplineReplacementType.Arcs => LocalizationManager.Instance.GetString("SplineReplacement_Arcs"),
        _ => type.ToString()
    };
}

public static class CsvDelimiterMapping
{
    private static readonly Dictionary<CsvDelimiterType, string> DelimiterMap = new()
    {
        { CsvDelimiterType.Comma, "," },
        { CsvDelimiterType.Semicolon, ";" },
        { CsvDelimiterType.Tab, "\t" },
        { CsvDelimiterType.Pipe, "|" }
    };

    public static string GetDisplayName(CsvDelimiterType type) => type switch
    {
        CsvDelimiterType.Comma => LocalizationManager.Instance.GetString("CsvDelimiter_Comma"),
        CsvDelimiterType.Semicolon => LocalizationManager.Instance.GetString("CsvDelimiter_Semicolon"),
        CsvDelimiterType.Tab => LocalizationManager.Instance.GetString("CsvDelimiter_Tab"),
        CsvDelimiterType.Pipe => LocalizationManager.Instance.GetString("CsvDelimiter_Pipe"),
        _ => type.ToString()
    };

    public static string GetDelimiter(CsvDelimiterType type) => DelimiterMap[type];
}