using netDxf.Header;

namespace FlatPatternExporter;

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