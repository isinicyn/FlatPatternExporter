namespace FlatPatternExporter.Utilities;

public static class VersionComparer
{
    public static bool IsNewerVersion(string currentVersion, string latestVersion)
    {
        var current = ParseVersion(currentVersion);
        var latest = ParseVersion(latestVersion);

        if (current == null || latest == null)
            return false;

        return latest > current;
    }

    public static Version? ParseVersion(string versionString)
    {
        if (string.IsNullOrWhiteSpace(versionString))
            return null;

        var cleaned = versionString.TrimStart('v').Trim();

        var parts = cleaned.Split(' ');
        var versionPart = parts[0];

        if (Version.TryParse(versionPart, out var version))
            return version;

        return null;
    }

    public static string ExtractVersionString(string fullVersionString)
    {
        if (string.IsNullOrWhiteSpace(fullVersionString))
            return string.Empty;

        var parts = fullVersionString.Split(' ');
        return parts[0];
    }
}
