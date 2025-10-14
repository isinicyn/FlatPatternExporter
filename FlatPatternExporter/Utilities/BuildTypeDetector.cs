using System.IO;
using FlatPatternExporter.Enums;

namespace FlatPatternExporter.Utilities;

public static class BuildTypeDetector
{
    private const string BuildTypeFileName = ".buildtype";

    public static BuildType DetectBuildType()
    {
        try
        {
            var appDirectory = AppContext.BaseDirectory;
            var buildTypeFilePath = Path.Combine(appDirectory, BuildTypeFileName);

            if (!File.Exists(buildTypeFilePath))
            {
                return BuildType.Portable;
            }

            var buildTypeContent = File.ReadAllText(buildTypeFilePath).Trim();

            return buildTypeContent switch
            {
                "Deploy" => BuildType.Deploy,
                "Portable" => BuildType.Portable,
                "FrameworkDependent" => BuildType.FrameworkDependent,
                _ => BuildType.Portable
            };
        }
        catch
        {
            return BuildType.Portable;
        }
    }
}
