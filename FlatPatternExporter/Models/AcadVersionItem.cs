using FlatPatternExporter.Enums;

namespace FlatPatternExporter.Models;

public class AcadVersionItem
{
    public string DisplayName { get; set; } = string.Empty;
    public AcadVersionType Value { get; set; }
    public override string ToString() => DisplayName;
}