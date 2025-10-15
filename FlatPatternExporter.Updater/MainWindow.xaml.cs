using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;

namespace FlatPatternExporter.Updater;

public partial class MainWindow : Window
{
    private readonly int _processId;
    private readonly string _updateFilePath;
    private readonly string _targetExecutablePath;
    private readonly string _logPath;
    private string _backupDirectory = string.Empty;

    public MainWindow(int processId, string updateFilesPath, string targetExecutablePath)
    {
        InitializeComponent();

        _processId = processId;
        _updateFilePath = updateFilesPath;
        _targetExecutablePath = targetExecutablePath;

        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var logsDirectory = Path.Combine(appDataPath, "FlatPatternExporter", "Logs");
        Directory.CreateDirectory(logsDirectory);

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
        _logPath = Path.Combine(logsDirectory, $"Updater_{timestamp}.log");

        RotateLogs(logsDirectory, maxLogFiles: 10);

        Log($"=== Updater Started ===");
        Log($"Process ID: {_processId}");
        Log($"Update files path: {_updateFilePath}");
        Log($"Target executable: {_targetExecutablePath}");

        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;

        this.Title = $"FlatPatternExporter - Updater";
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        Log("MainWindow_Closing event fired");
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        Log("MainWindow_Closed event fired");
    }

    private void Log(string message)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var logMessage = $"[{timestamp}] {message}\n";
            File.AppendAllText(_logPath, logMessage);

            // Force flush to disk immediately
            using (var fs = new FileStream(_logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            {
                fs.Flush(true);
            }
        }
        catch
        {
        }
    }

    private void RotateLogs(string logsDirectory, int maxLogFiles)
    {
        try
        {
            var logFiles = Directory.GetFiles(logsDirectory, "Updater_*.log")
                .Select(f => new FileInfo(f))
                .OrderByDescending(f => f.CreationTime)
                .ToList();

            if (logFiles.Count > maxLogFiles)
            {
                var filesToDelete = logFiles.Skip(maxLogFiles);
                foreach (var file in filesToDelete)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                    }
                }
            }
        }
        catch
        {
        }
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Log("MainWindow_Loaded started");
        SafeUpdateDetailsText($"Log file: {_logPath}\n\nPreparing for update...");

        try
        {
            Log("Calling PerformUpdateAsync");
            await PerformUpdateAsync();
            Log("PerformUpdateAsync completed successfully");
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
            Log($"STACK TRACE: {ex.StackTrace}");

            UpdateStatus($"Update error: {ex.Message}", 0);
            SafeUpdateProgressText($"Error");
            SafeUpdateDetailsText($"Update error:\n{ex.Message}\n\nLog file:\n{_logPath}");

            MessageBox.Show(
                this,
                $"Update error:\n\n{ex.Message}\n\n{ex.StackTrace}\n\nLog file: {_logPath}",
                "Update Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            await Task.Delay(5000);
            Application.Current.Shutdown(1);
        }
    }

    private async Task PerformUpdateAsync()
    {
        var updateSuccessful = false;

        try
        {
            Log("Step 1: Waiting for process to exit");
            UpdateStatus("Waiting for application to exit...", 10);

            Log("About to await WaitForProcessToExitAsync");
            await WaitForProcessToExitAsync(_processId);
            Log("WaitForProcessToExitAsync returned");
            Log("Step 1 completed");

            Log("Step 2: Creating backup");
            UpdateStatus("Creating backup...", 30);
            Log("About to create backup");
            CreateBackup(_targetExecutablePath);
            Log("Backup created");
            Log("Step 2 completed");

            Log("Step 3: Replacing application files");
            UpdateStatus("Replacing application files...", 60);
            Log("About to replace files");

            try
            {
                ReplaceExecutable(_updateFilePath, _targetExecutablePath);
                Log("Files replaced successfully");
                Log("Step 3 completed");
                updateSuccessful = true;
            }
            catch (Exception ex)
            {
                Log($"ERROR during file replacement: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");

                Log("Step 4: Rolling back from backup");
                UpdateStatus("Update failed. Rolling back...", 50);

                try
                {
                    RestoreFromBackup(_targetExecutablePath);
                    Log("Rollback completed successfully");

                    UpdateStatus("Update failed. System restored to previous state.", 0);
                    SafeUpdateProgressText("Update failed - Rolled back");
                    SafeUpdateDetailsText($"Update failed and rolled back:\n{ex.Message}\n\nYour application has been restored to the previous state.\n\nLog file:\n{_logPath}");

                    MessageBox.Show(
                        this,
                        $"Update failed and rolled back:\n\n{ex.Message}\n\nYour application has been restored to the previous state.\n\nLog file: {_logPath}",
                        "Update Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                catch (Exception rollbackEx)
                {
                    Log($"CRITICAL: Rollback failed: {rollbackEx.Message}");
                    Log($"Stack trace: {rollbackEx.StackTrace}");

                    UpdateStatus("CRITICAL: Update and rollback failed!", 0);
                    SafeUpdateProgressText("Critical Error");
                    SafeUpdateDetailsText($"Update failed and rollback also failed!\n\nOriginal error:\n{ex.Message}\n\nRollback error:\n{rollbackEx.Message}\n\nBackup location:\n{_backupDirectory}\n\nLog file:\n{_logPath}");

                    MessageBox.Show(
                        this,
                        $"CRITICAL ERROR:\n\nUpdate failed:\n{ex.Message}\n\nRollback also failed:\n{rollbackEx.Message}\n\nPlease manually restore from backup:\n{_backupDirectory}\n\nLog file: {_logPath}",
                        "Critical Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }

                await Task.Delay(5000);
                Application.Current.Shutdown(1);
                return;
            }

            if (updateSuccessful)
            {
                Log("Step 4: Cleaning up backup");
                UpdateStatus("Cleaning up...", 90);
                DeleteBackup();
                Log("Backup cleanup completed");

                Log("All steps completed successfully");
                UpdateStatus("Update completed successfully!", 100);
                SafeUpdateProgressText("Ready to restart");
                SafeUpdateDetailsText("All steps completed successfully!\nUpdate finished.\n\nClick OK to restart the application.");

                Log("Enabling close button");
                Dispatcher.Invoke(() =>
                {
                    CloseButton.IsEnabled = true;
                });
            }
        }
        catch (Exception ex)
        {
            Log($"EXCEPTION in PerformUpdateAsync: {ex.Message}");
            Log($"Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private async Task WaitForProcessToExitAsync(int processId)
    {
        try
        {
            Log($"Looking for process with ID {processId}");
            var process = Process.GetProcessById(processId);
            Log($"Process found: {process.ProcessName}");
            SafeUpdateProgressText($"Waiting for process to exit...");
            SafeUpdateDetailsText($"Process: {process.ProcessName}\nPID: {processId}");
            await process.WaitForExitAsync();
            Log("Process has exited");
            SafeUpdateProgressText("Process exited");
            SafeUpdateDetailsText($"Process {process.ProcessName} successfully exited");
        }
        catch (ArgumentException ex)
        {
            Log($"Process not found (already exited): {ex.Message}");
            SafeUpdateProgressText("Process already exited");
            SafeUpdateDetailsText($"Process with PID {processId} already exited");
        }
        Log("WaitForProcessToExitAsync completed");
    }

    private void SafeUpdateProgressText(string message)
    {
        try
        {
            Log($"SafeUpdateProgressText: {message}");
            Dispatcher.Invoke(() =>
            {
                ProgressTextBlock.Text = message;
            });
            Log($"SafeUpdateProgressText completed");
        }
        catch (Exception ex)
        {
            Log($"ERROR in SafeUpdateProgressText: {ex.Message}");
        }
    }

    private void SafeUpdateDetailsText(string message)
    {
        try
        {
            Dispatcher.Invoke(() =>
            {
                DetailsTextBlock.Text = message;
            });
        }
        catch (Exception ex)
        {
            Log($"ERROR in SafeUpdateDetailsText: {ex.Message}");
        }
    }

    private static bool IsBackupPath(string path)
    {
        var segments = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return segments.Any(segment => segment.StartsWith("backup_", StringComparison.OrdinalIgnoreCase));
    }

    private void CreateBackup(string targetExecutablePath)
    {
        var targetDirectory = Path.GetDirectoryName(targetExecutablePath);
        if (string.IsNullOrEmpty(targetDirectory))
        {
            throw new InvalidOperationException("Target directory is null or empty");
        }

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _backupDirectory = Path.Combine(targetDirectory, $"backup_{timestamp}");

        Log($"Creating backup directory at: {_backupDirectory}");
        SafeUpdateProgressText($"Creating backup...");
        SafeUpdateDetailsText($"Backup directory: {Path.GetFileName(_backupDirectory)}");

        if (Directory.Exists(_backupDirectory))
        {
            Log($"Deleting existing backup directory");
            Directory.Delete(_backupDirectory, recursive: true);
        }

        Directory.CreateDirectory(_backupDirectory);

        var filesToBackup = Directory.GetFiles(targetDirectory, "*.*", SearchOption.AllDirectories)
            .Where(file => !IsBackupPath(file))
            .ToList();

        Log($"Found {filesToBackup.Count} files to backup");

        foreach (var sourceFile in filesToBackup)
        {
            var relativePath = Path.GetRelativePath(targetDirectory, sourceFile);
            var backupFile = Path.Combine(_backupDirectory, relativePath);
            var backupFileDirectory = Path.GetDirectoryName(backupFile);

            if (!string.IsNullOrEmpty(backupFileDirectory) && !Directory.Exists(backupFileDirectory))
            {
                Directory.CreateDirectory(backupFileDirectory);
            }

            Log($"Backing up: {relativePath}");
            SafeUpdateDetailsText($"Backing up: {relativePath}");
            File.Copy(sourceFile, backupFile, overwrite: true);
        }

        Log($"Backup created successfully. Total files: {filesToBackup.Count}");
        SafeUpdateProgressText($"Backup created");
        SafeUpdateDetailsText($"Backup successfully created:\n{Path.GetFileName(_backupDirectory)}\nTotal files: {filesToBackup.Count}");
    }

    private void ReplaceExecutable(string updateFilesPath, string targetExecutablePath)
    {
        Log($"Replacing application files from: {updateFilesPath}");
        SafeUpdateProgressText("Replacing application files...");

        var targetDirectory = Path.GetDirectoryName(targetExecutablePath);
        if (string.IsNullOrEmpty(targetDirectory))
        {
            throw new InvalidOperationException("Target directory is null or empty");
        }

        var updateFiles = Directory.GetFiles(updateFilesPath, "*.*", SearchOption.AllDirectories);
        Log($"Found {updateFiles.Length} files to update");

        var retryCount = 0;
        const int maxRetries = 5;

        foreach (var sourceFile in updateFiles)
        {
            var relativePath = Path.GetRelativePath(updateFilesPath, sourceFile);
            var targetFile = Path.Combine(targetDirectory, relativePath);
            var targetFileDirectory = Path.GetDirectoryName(targetFile);

            if (!string.IsNullOrEmpty(targetFileDirectory) && !Directory.Exists(targetFileDirectory))
            {
                Directory.CreateDirectory(targetFileDirectory);
            }

            retryCount = 0;
            var copied = false;

            while (retryCount < maxRetries && !copied)
            {
                try
                {
                    Log($"Copying: {relativePath}");
                    SafeUpdateDetailsText($"Copying: {relativePath}");
                    File.Copy(sourceFile, targetFile, overwrite: true);
                    Log($"File copied successfully: {relativePath}");
                    copied = true;
                }
                catch (IOException ex)
                {
                    retryCount++;
                    if (retryCount >= maxRetries)
                    {
                        var errorMsg = $"Failed to replace file '{relativePath}' after {maxRetries} attempts: {ex.Message}";
                        Log($"ERROR: {errorMsg}");
                        throw new IOException(errorMsg, ex);
                    }

                    Log($"File locked, retry {retryCount}/{maxRetries}: {ex.Message}");
                    SafeUpdateProgressText($"Attempt {retryCount}/{maxRetries}...");
                    SafeUpdateDetailsText($"File locked, retrying...\nFile: {relativePath}\nAttempt {retryCount} of {maxRetries}");
                    Thread.Sleep(1000);
                }
            }
        }

        Log("All files replaced successfully");
        SafeUpdateProgressText("Files successfully replaced");
        SafeUpdateDetailsText($"Application files successfully updated\nTotal files: {updateFiles.Length}");
    }

    private void DeleteBackup()
    {
        if (string.IsNullOrEmpty(_backupDirectory) || !Directory.Exists(_backupDirectory))
        {
            Log("No backup directory to delete");
            return;
        }

        Log($"Deleting backup directory: {_backupDirectory}");
        SafeUpdateProgressText("Cleaning up backup...");
        SafeUpdateDetailsText($"Removing backup directory:\n{Path.GetFileName(_backupDirectory)}");

        try
        {
            Directory.Delete(_backupDirectory, recursive: true);
            Log("Backup directory deleted successfully");
            SafeUpdateProgressText("Backup cleaned up");
        }
        catch (Exception ex)
        {
            Log($"ERROR deleting backup: {ex.Message}");
            SafeUpdateProgressText("Failed to clean up backup");
        }
    }

    private void RestoreFromBackup(string targetExecutablePath)
    {
        if (string.IsNullOrEmpty(_backupDirectory) || !Directory.Exists(_backupDirectory))
        {
            Log("No backup directory to restore from");
            throw new InvalidOperationException("Backup directory not found. Cannot restore.");
        }

        var targetDirectory = Path.GetDirectoryName(targetExecutablePath);
        if (string.IsNullOrEmpty(targetDirectory))
        {
            throw new InvalidOperationException("Target directory is null or empty");
        }

        Log($"Restoring from backup: {_backupDirectory}");
        SafeUpdateProgressText("Cleaning target directory...");
        SafeUpdateDetailsText($"Rolling back changes...\nStep 1: Removing all files");

        Log("Step 1: Removing all files from target directory (except backup)");
        var filesToDelete = Directory.GetFiles(targetDirectory, "*.*", SearchOption.AllDirectories)
            .Where(file => !IsBackupPath(file))
            .ToList();

        Log($"Found {filesToDelete.Count} files to delete");

        foreach (var file in filesToDelete)
        {
            try
            {
                var relativePath = Path.GetRelativePath(targetDirectory, file);
                Log($"Deleting: {relativePath}");
                SafeUpdateDetailsText($"Removing: {relativePath}");
                File.Delete(file);
            }
            catch (Exception ex)
            {
                Log($"WARNING: Failed to delete file {file}: {ex.Message}");
            }
        }

        Log("Step 2: Removing empty directories");
        var directoriesToDelete = Directory.GetDirectories(targetDirectory, "*", SearchOption.AllDirectories)
            .Where(dir => !IsBackupPath(dir))
            .OrderByDescending(dir => dir.Length)
            .ToList();

        foreach (var directory in directoriesToDelete)
        {
            try
            {
                if (Directory.Exists(directory) && !Directory.EnumerateFileSystemEntries(directory).Any())
                {
                    var relativePath = Path.GetRelativePath(targetDirectory, directory);
                    Log($"Deleting empty directory: {relativePath}");
                    Directory.Delete(directory);
                }
            }
            catch (Exception ex)
            {
                Log($"WARNING: Failed to delete directory {directory}: {ex.Message}");
            }
        }

        Log("Step 3: Restoring files from backup");
        SafeUpdateProgressText("Restoring files from backup...");
        SafeUpdateDetailsText($"Step 2: Restoring files\nBackup: {Path.GetFileName(_backupDirectory)}");

        var backupFiles = Directory.GetFiles(_backupDirectory, "*.*", SearchOption.AllDirectories);
        Log($"Found {backupFiles.Length} files to restore");

        foreach (var backupFile in backupFiles)
        {
            var relativePath = Path.GetRelativePath(_backupDirectory, backupFile);
            var targetFile = Path.Combine(targetDirectory, relativePath);
            var targetFileDirectory = Path.GetDirectoryName(targetFile);

            if (!string.IsNullOrEmpty(targetFileDirectory) && !Directory.Exists(targetFileDirectory))
            {
                Directory.CreateDirectory(targetFileDirectory);
            }

            Log($"Restoring: {relativePath}");
            SafeUpdateDetailsText($"Restoring: {relativePath}");
            File.Copy(backupFile, targetFile, overwrite: true);
        }

        Log($"Restore completed. Files deleted: {filesToDelete.Count}, Files restored: {backupFiles.Length}");
        SafeUpdateProgressText("Restore completed");
        SafeUpdateDetailsText($"Rollback completed successfully!\n\nFiles removed: {filesToDelete.Count}\nFiles restored: {backupFiles.Length}");
    }

    private void RestartApplication(string targetExecutablePath)
    {
        Log($"Restarting application: {targetExecutablePath}");
        SafeUpdateProgressText("Launching updated application...");
        SafeUpdateDetailsText($"Launching: {Path.GetFileName(targetExecutablePath)}");

        var workingDir = Path.GetDirectoryName(targetExecutablePath);
        Log($"Working directory: {workingDir}");

        var startInfo = new ProcessStartInfo
        {
            FileName = targetExecutablePath,
            UseShellExecute = true,
            WorkingDirectory = workingDir
        };

        var process = Process.Start(startInfo);
        Log($"Application started with PID: {process?.Id ?? -1}");
        SafeUpdateProgressText("Application successfully launched");
        SafeUpdateDetailsText($"Application successfully launched\nPID: {process?.Id ?? -1}");
    }

    private void UpdateStatus(string message, double progress)
    {
        try
        {
            Log($"UpdateStatus: {message} ({progress}%)");
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = message;
                UpdateProgressBar.Value = progress;
            });
        }
        catch (Exception ex)
        {
            Log($"ERROR in UpdateStatus: {ex.Message}");
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Log("Close button clicked");

        try
        {
            Log("Restarting application");
            RestartApplication(_targetExecutablePath);
            Log("Application restarted");

            Log("Shutting down updater");
            Application.Current.Shutdown(0);
        }
        catch (Exception ex)
        {
            Log($"ERROR restarting application: {ex.Message}");
            MessageBox.Show(
                this,
                $"Failed to restart application:\n\n{ex.Message}",
                "Restart Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            Log("Shutting down updater despite error");
            Application.Current.Shutdown(1);
        }
    }
}
