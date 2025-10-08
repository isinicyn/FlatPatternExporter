using System.Windows;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.UI.Windows;

public partial class UpdateWindow : Window
{
    private readonly UpdateManager _updateManager;
    private readonly ReleaseInfo _releaseInfo;
    private readonly string _currentVersion;
    private string? _downloadedFilePath;

    public UpdateWindow(UpdateCheckResult updateResult)
    {
        InitializeComponent();

        _updateManager = new UpdateManager();
        _releaseInfo = updateResult.ReleaseInfo!;
        _currentVersion = updateResult.CurrentVersion;

        InitializeUI();
    }

    private void InitializeUI()
    {
        CurrentVersionTextBlock.Text = _currentVersion;
        NewVersionTextBlock.Text = _releaseInfo.Version;
        ReleaseNotesTextBlock.Text = string.IsNullOrWhiteSpace(_releaseInfo.Body)
            ? LocalizationManager.Instance.GetString("Text_NoReleaseNotes")
            : _releaseInfo.Body;
    }

    private async void InstallButton_Click(object sender, RoutedEventArgs e)
    {
        InstallButton.IsEnabled = false;
        CancelButton.IsEnabled = false;
        ProgressGroupBox.Visibility = Visibility.Visible;
        StatusTextBlock.Text = LocalizationManager.Instance.GetString("Status_DownloadingUpdate");

        var progress = new Progress<double>(value =>
        {
            DownloadProgressBar.Value = value;
            ProgressTextBlock.Text = $"{value:F1}%";
        });

        _downloadedFilePath = await _updateManager.DownloadUpdateAsync(_releaseInfo, progress);

        if (_downloadedFilePath == null)
        {
            StatusTextBlock.Text = LocalizationManager.Instance.GetString("Error_DownloadFailed");
            InstallButton.IsEnabled = true;
            CancelButton.IsEnabled = true;
            return;
        }

        StatusTextBlock.Text = LocalizationManager.Instance.GetString("Status_LaunchingUpdater");

        try
        {
            _updateManager.LaunchUpdater(_downloadedFilePath);
            System.Windows.Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"{LocalizationManager.Instance.GetString("Error_UpdateFailed")}: {ex.Message}";
            InstallButton.IsEnabled = true;
            CancelButton.IsEnabled = true;
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
