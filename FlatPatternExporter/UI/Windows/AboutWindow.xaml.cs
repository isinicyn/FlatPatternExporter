using System.Windows;
using FlatPatternExporter.Services;
using WpfToolkit;
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

        /// <summary>
        /// Fixes visual artifacts that occur when using WindowChrome with SizeToContent.
        /// Forces window size update by toggling SizeToContent mode.
        /// </summary>
        private void RefreshWindowSize(object sender, EventArgs e)
        {
            SizeToContent = SizeToContent.Manual;
            SizeToContent = SizeToContent.WidthAndHeight;
        }

        private void SetVersion()
        {
            var version = VersionInfoService.GetApplicationVersion();
            var lastUpdate = VersionInfoService.GetLastCommitDate();
            VersionTextBlock.Text = $"{version} ({lastUpdate})";
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CopyVersionToClipboard_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(VersionTextBlock.Text);
            var message = LocalizationManager.Instance.GetString("Text_Copied");
            PopupNotificationService.ShowNotification(CopyButton, message);
        }

        private void EmailTextBlock_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Clipboard.SetText(EmailTextBlock.Text);
            var message = LocalizationManager.Instance.GetString("Text_Copied");
            PopupNotificationService.ShowNotification(EmailTextBlock, message);
        }
    }
}
