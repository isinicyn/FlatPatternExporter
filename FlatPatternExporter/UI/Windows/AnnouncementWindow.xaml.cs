using System.Diagnostics;
using System.Windows;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.UI.Windows;

public partial class AnnouncementWindow : Window
{
    private readonly string _pageUrl;

    private AnnouncementWindow(ReleaseInfo? release)
    {
        InitializeComponent();

        var (header, subtitle, button) = release != null
            ? ("Announcement_HeaderReleased", "Announcement_SubtitleReleased", "Button_OpenDownloadPage")
            : ("Announcement_HeaderUpcoming", "Announcement_SubtitleUpcoming", "Button_OpenProjectPage");

        var localization = LocalizationManager.Instance;
        HeaderTextBlock.Text = localization.GetString(header);
        SubtitleTextBlock.Text = localization.GetString(subtitle);
        OpenPageButton.Content = localization.GetString(button);
        _pageUrl = release?.HtmlUrl ?? GitHubUpdateService.ReleasesUrl;
    }

    /// <summary>
    /// Shows the announcement of the released version 4 or, without a release, of the upcoming one
    /// </summary>
    public static void ShowFor(Window owner, ReleaseInfo? release)
    {
        if (System.Windows.Application.Current.Windows.OfType<AnnouncementWindow>().FirstOrDefault() is { } current)
        {
            current.Activate();
            return;
        }

        new AnnouncementWindow(release) { Owner = owner }.Show();
    }

    private void OpenPageButton_Click(object sender, RoutedEventArgs e)
    {
        Process.Start(new ProcessStartInfo(_pageUrl) { UseShellExecute = true });
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
