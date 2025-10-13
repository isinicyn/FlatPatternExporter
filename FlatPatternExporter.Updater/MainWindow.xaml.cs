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

    public MainWindow(int processId, string updateFilePath, string targetExecutablePath)
    {
        InitializeComponent();

        _processId = processId;
        _updateFilePath = updateFilePath;
        _targetExecutablePath = targetExecutablePath;

        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var logsDirectory = Path.Combine(appDataPath, "FlatPatternExporter", "Logs");
        Directory.CreateDirectory(logsDirectory);

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
        _logPath = Path.Combine(logsDirectory, $"Updater_{timestamp}.log");

        RotateLogs(logsDirectory, maxLogFiles: 10);

        Log($"=== Updater Started ===");
        Log($"Process ID: {_processId}");
        Log($"Update file: {_updateFilePath}");
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

            Log("Step 3: Replacing executable");
            UpdateStatus("Replacing application file...", 60);
            Log("About to replace executable");
            ReplaceExecutable(_updateFilePath, _targetExecutablePath);
            Log("Executable replaced");
            Log("Step 3 completed");

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

    private void CreateBackup(string targetExecutablePath)
    {
        var backupPath = targetExecutablePath + ".backup";
        Log($"Creating backup at: {backupPath}");
        SafeUpdateProgressText($"Creating backup...");
        SafeUpdateDetailsText($"File: {Path.GetFileName(targetExecutablePath)}\nBackup: {Path.GetFileName(backupPath)}");

        if (File.Exists(backupPath))
        {
            Log($"Deleting existing backup");
            File.Delete(backupPath);
        }

        Log($"Copying {targetExecutablePath} to {backupPath}");
        File.Copy(targetExecutablePath, backupPath, overwrite: true);
        Log($"Backup created successfully");
        SafeUpdateProgressText($"Backup created");
        SafeUpdateDetailsText($"Backup successfully created:\n{Path.GetFileName(backupPath)}");
    }

    private void ReplaceExecutable(string updateFilePath, string targetExecutablePath)
    {
        Log($"Replacing executable: {updateFilePath} -> {targetExecutablePath}");
        SafeUpdateProgressText("Replacing executable...");
        SafeUpdateDetailsText($"Source: {Path.GetFileName(updateFilePath)}\nDestination: {Path.GetFileName(targetExecutablePath)}");

        var retryCount = 0;
        const int maxRetries = 5;

        while (retryCount < maxRetries)
        {
            try
            {
                Log($"Attempt {retryCount + 1} to copy file");
                File.Copy(updateFilePath, targetExecutablePath, overwrite: true);
                Log("File copied successfully");
                SafeUpdateProgressText("File successfully replaced");
                SafeUpdateDetailsText($"Application file successfully updated:\n{Path.GetFileName(targetExecutablePath)}");
                return;
            }
            catch (IOException ex) when (retryCount < maxRetries - 1)
            {
                retryCount++;
                Log($"File locked, retry {retryCount}/{maxRetries}: {ex.Message}");
                SafeUpdateProgressText($"Attempt {retryCount}/{maxRetries}...");
                SafeUpdateDetailsText($"File locked, retrying...\nAttempt {retryCount} of {maxRetries}");
                Thread.Sleep(1000);
            }
        }

        var errorMsg = $"Failed to replace file after {maxRetries} attempts";
        Log($"ERROR: {errorMsg}");
        throw new IOException(errorMsg);
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
