using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using FlatPatternExporter.Enums;
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
            var buildType = BuildTypeDetector.DetectBuildType();
            var buildSuffix = BuildTypeMapping.GetArchiveSuffix(buildType);
            var version = release.Version;

            var updaterAssetName = $"FlatPatternExporter.Updater-v{version}-x64.zip";
            var mainAppAssetName = $"FlatPatternExporter-v{version}-x64-{buildSuffix}.zip";

            var updaterAsset = release.Assets.FirstOrDefault(a =>
                a.Name.Equals(updaterAssetName, StringComparison.OrdinalIgnoreCase));
            var mainAppAsset = release.Assets.FirstOrDefault(a =>
                a.Name.Equals(mainAppAssetName, StringComparison.OrdinalIgnoreCase));

            if (updaterAsset == null || mainAppAsset == null)
            {
                var missingAsset = updaterAsset == null ? updaterAssetName : mainAppAssetName;
                throw new FileNotFoundException($"Required archive not found in release: {missingAsset}");
            }

            var tempPath = Path.Combine(Path.GetTempPath(), "FlatPatternExporter_Update");

            if (Directory.Exists(tempPath))
            {
                try
                {
                    Directory.Delete(tempPath, recursive: true);
                }
                catch
                {
                }
            }

            Directory.CreateDirectory(tempPath);

            var updaterZipPath = Path.Combine(tempPath, updaterAssetName);
            var mainAppZipPath = Path.Combine(tempPath, mainAppAssetName);

            var updaterProgress = new Progress<double>(p => progress?.Report(p * 0.3));
            var success = await _githubService.DownloadReleaseAssetAsync(updaterAsset.DownloadUrl, updaterZipPath, updaterProgress);

            if (!success)
            {
                CleanupTempDirectory(tempPath);
                return null;
            }

            var mainAppProgress = new Progress<double>(p => progress?.Report(30 + p * 0.7));
            success = await _githubService.DownloadReleaseAssetAsync(mainAppAsset.DownloadUrl, mainAppZipPath, mainAppProgress);

            if (!success)
            {
                CleanupTempDirectory(tempPath);
                return null;
            }

            var updaterExtractPath = Path.Combine(tempPath, "updater");
            var mainAppExtractPath = Path.Combine(tempPath, "app");

            await Task.Run(() => ZipFile.ExtractToDirectory(updaterZipPath, updaterExtractPath));
            await Task.Run(() => ZipFile.ExtractToDirectory(mainAppZipPath, mainAppExtractPath));

            File.Delete(updaterZipPath);
            File.Delete(mainAppZipPath);

            var updaterExePath = Path.Combine(updaterExtractPath, "FlatPatternExporter.Updater.exe");
            return File.Exists(updaterExePath) ? mainAppExtractPath : null;
        }
        catch
        {
            var tempPath = Path.Combine(Path.GetTempPath(), "FlatPatternExporter_Update");
            CleanupTempDirectory(tempPath);
            return null;
        }
    }

    public void LaunchUpdater(string updateFilesPath)
    {
        var currentExecutable = Environment.ProcessPath ?? string.Empty;
        var tempPath = Path.GetDirectoryName(updateFilesPath) ?? string.Empty;
        var updaterPath = Path.Combine(tempPath, "updater", "FlatPatternExporter.Updater.exe");

        if (!File.Exists(updaterPath))
        {
            throw new FileNotFoundException("Updater executable not found", updaterPath);
        }

        var currentProcessId = Environment.ProcessId;
        var arguments = $"\"{currentProcessId}\" \"{updateFilesPath}\" \"{currentExecutable}\"";

        var startInfo = new ProcessStartInfo
        {
            FileName = updaterPath,
            Arguments = arguments,
            UseShellExecute = true,
            WorkingDirectory = tempPath
        };

        Process.Start(startInfo);
    }

    private static void CleanupTempDirectory(string tempPath)
    {
        if (Directory.Exists(tempPath))
        {
            try
            {
                Directory.Delete(tempPath, recursive: true);
            }
            catch
            {
            }
        }
    }
}
