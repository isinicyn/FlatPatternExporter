using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
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
            ApplyGrayscaleEffect();
        }

        private void ApplyGrayscaleEffect()
        {
            var originalSource = GrayscaleImage.Source as BitmapSource;
            if (originalSource != null)
            {
                var grayBitmap = new FormatConvertedBitmap();
                grayBitmap.BeginInit();
                grayBitmap.Source = originalSource;
                grayBitmap.DestinationFormat = System.Windows.Media.PixelFormats.Gray8;
                grayBitmap.EndInit();
                GrayscaleImage.Source = grayBitmap;
            }
        }

        private void ImageBorder_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var fadeIn = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromMilliseconds(1000),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
            };
            ColorImage.BeginAnimation(OpacityProperty, fadeIn);
        }

        private void ImageBorder_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var fadeOut = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromMilliseconds(300),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
            };
            ColorImage.BeginAnimation(OpacityProperty, fadeOut);
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
