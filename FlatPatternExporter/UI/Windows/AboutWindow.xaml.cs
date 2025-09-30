using System.Windows;
using FlatPatternExporter.Services;
using Clipboard = System.Windows.Clipboard;

namespace FlatPatternExporter.UI.Windows
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            SetVersion();
            SetLastUpdateDate();
        }

        private void SetVersion()
        {
            VersionTextBlock.Text = LocalizationManager.Instance.GetString("About_ProgramVersion") + " " + VersionInfoService.GetApplicationVersion();
        }

        // Set the last update date according to commit
        private void SetLastUpdateDate()
        {
            LastUpdateTextBlock.Text = LocalizationManager.Instance.GetString("About_LastUpdate") + " " + VersionInfoService.GetLastCommitDate();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CopyVersionToClipboard_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(VersionTextBlock.Text);
        }
    }
}