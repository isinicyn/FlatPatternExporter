using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace FlatPatternExporter.Services;

public class VersionInfoService
{
    public static string GetApplicationVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        return informationalVersion ?? assembly.GetName().Version?.ToString() ?? "Version unknown";
    }

    public static string GetLastCommitDate()
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var commitDateAttribute = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "CommitDate");

            if (commitDateAttribute != null && !string.IsNullOrWhiteSpace(commitDateAttribute.Value))
            {
                return commitDateAttribute.Value;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error retrieving commit date from assembly metadata: {ex.Message}");
        }

        return "Date not determined";
    }
}