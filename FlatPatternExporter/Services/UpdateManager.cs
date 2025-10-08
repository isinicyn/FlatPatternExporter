using System.Diagnostics;
using System.IO;
using FlatPatternExporter.Models;
using FlatPatternExporter.Utilities;

namespace FlatPatternExporter.Services;

public class UpdateManager
{
    private readonly GitHubUpdateService _githubService;

    public UpdateManager()
    {
        _githubService = new GitHubUpdateService();
    }

    public async Task<UpdateCheckResult> CheckForUpdatesAsync()
    {
        try
        {
            var currentVersionString = VersionInfoService.GetApplicationVersion();
            var currentVersion = VersionComparer.ExtractVersionString(currentVersionString);

            var latestRelease = await _githubService.GetLatestReleaseAsync();

            if (latestRelease == null)
            {
                string errorMessage;

                if (_githubService.LastError == "RATE_LIMIT_EXCEEDED")
                {
                    errorMessage = LocalizationManager.Instance.GetString("Error_RateLimitExceeded");
                }
                else
                {
                    errorMessage = "Failed to fetch release information from GitHub";
                    if (!string.IsNullOrEmpty(_githubService.LastError))
                    {
                        errorMessage += $"\n\nDetails:\n{_githubService.LastError}";
                    }
                }

                return new UpdateCheckResult
                {
                    ErrorMessage = errorMessage,
                    CurrentVersion = currentVersion
                };
            }

            var latestVersion = latestRelease.Version;
            var isUpdateAvailable = VersionComparer.IsNewerVersion(currentVersion, latestVersion);

            return new UpdateCheckResult
            {
                IsUpdateAvailable = isUpdateAvailable,
                CurrentVersion = currentVersion,
                LatestVersion = latestVersion,
                ReleaseInfo = latestRelease
            };
        }
        catch (Exception ex)
        {
            return new UpdateCheckResult
            {
                ErrorMessage = $"Error checking for updates: {ex.Message}\n\nStack trace:\n{ex.StackTrace}"
            };
        }
    }

    public async Task<string?> DownloadUpdateAsync(ReleaseInfo release, IProgress<double>? progress = null)
    {
        try
        {
            var asset = release.Assets.FirstOrDefault(a => a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase));

            if (asset == null)
                return null;

            var tempPath = Path.Combine(Path.GetTempPath(), "FlatPatternExporter_Update");
            Directory.CreateDirectory(tempPath);

            var downloadPath = Path.Combine(tempPath, asset.Name);

            var success = await _githubService.DownloadReleaseAssetAsync(asset.DownloadUrl, downloadPath, progress);

            return success ? downloadPath : null;
        }
        catch
        {
            return null;
        }
    }

    public void LaunchUpdater(string updateFilePath)
    {
        var currentExecutable = Environment.ProcessPath ?? string.Empty;
        var currentDirectory = Path.GetDirectoryName(currentExecutable) ?? string.Empty;
        var updaterPath = Path.Combine(currentDirectory, "FlatPatternExporter.Updater.exe");

        if (!File.Exists(updaterPath))
        {
            throw new FileNotFoundException("Updater executable not found", updaterPath);
        }

        var currentProcessId = Environment.ProcessId;
        var arguments = $"\"{currentProcessId}\" \"{updateFilePath}\" \"{currentExecutable}\"";

        var startInfo = new ProcessStartInfo
        {
            FileName = updaterPath,
            Arguments = arguments,
            UseShellExecute = true,
            WorkingDirectory = currentDirectory
        };

        Process.Start(startInfo);
    }
}
