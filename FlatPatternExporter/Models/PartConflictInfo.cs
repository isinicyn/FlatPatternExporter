namespace FlatPatternExporter.Models;

public class PartConflictInfo
{
    public string PartNumber { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string ModelState { get; set; } = string.Empty;

    public string UniqueId => $"{FileName}|{ModelState}";
}