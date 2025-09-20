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
            VersionTextBlock.Text = "Версия программы: " + VersionInfoService.GetApplicationVersion();
        }

        // Установка даты последнего обновления согласно коммиту
        private void SetLastUpdateDate()
        {
            LastUpdateTextBlock.Text = "Последнее обновление: " + VersionInfoService.GetLastCommitDate();
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