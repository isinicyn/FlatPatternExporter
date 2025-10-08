namespace FlatPatternExporter.Models;

public class UpdateCheckResult
{
    public bool IsUpdateAvailable { get; set; }
    public string CurrentVersion { get; set; } = string.Empty;
    public string LatestVersion { get; set; } = string.Empty;
    public ReleaseInfo? ReleaseInfo { get; set; }
    public string? ErrorMessage { get; set; }
    public bool Success => ErrorMessage == null;
}
