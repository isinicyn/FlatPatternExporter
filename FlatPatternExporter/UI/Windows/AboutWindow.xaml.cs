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
        }

        private void SetVersion()
        {
            var version = VersionInfoService.GetApplicationVersion();
            var lastUpdate = VersionInfoService.GetLastCommitDate();
            VersionTextBlock.Text = $"{LocalizationManager.Instance.GetString("About_ProgramVersion")} {version} ({lastUpdate})";
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