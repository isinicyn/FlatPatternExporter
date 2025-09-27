using FlatPatternExporter.Enums;

namespace FlatPatternExporter.Models;

public class SplineReplacementItem
{
    public string DisplayName { get; set; } = string.Empty;
    public SplineReplacementType Value { get; set; }
    public override string ToString() => DisplayName;
}