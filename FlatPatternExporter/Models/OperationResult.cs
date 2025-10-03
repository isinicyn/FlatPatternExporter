using FlatPatternExporter.Enums;

namespace FlatPatternExporter.Models;

public class OperationResult
{
    public int ProcessedCount { get; set; }
    public int SkippedCount { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public bool WasCancelled { get; set; }
    public ProcessingMethod? ProcessingMethod { get; set; }
    public List<string> Errors { get; set; } = [];
}