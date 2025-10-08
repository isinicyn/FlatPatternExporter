using System.IO;
using System.Windows;

namespace FlatPatternExporter.Updater;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Length != 3)
        {
            MessageBox.Show(
                "Usage: FlatPatternExporter.Updater.exe <ProcessId> <UpdateFilePath> <TargetExecutablePath>",
                "Invalid Arguments",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        if (!int.TryParse(e.Args[0], out var processId))
        {
            MessageBox.Show(
                "Error: Invalid process ID",
                "Invalid Arguments",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        var updateFilePath = e.Args[1];
        var targetExecutablePath = e.Args[2];

        if (!File.Exists(updateFilePath))
        {
            MessageBox.Show(
                $"Error: Update file not found:\n{updateFilePath}",
                "File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        if (!File.Exists(targetExecutablePath))
        {
            MessageBox.Show(
                $"Error: Target executable not found:\n{targetExecutablePath}",
                "File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        var mainWindow = new MainWindow(processId, updateFilePath, targetExecutablePath);
        mainWindow.Show();
    }
}
