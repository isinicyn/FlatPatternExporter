using System.Windows;
using System.Windows.Controls;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using System.ComponentModel;

namespace FlatPatternExporter.UI.Controls
{
    public partial class LayerSettingControl : System.Windows.Controls.UserControl, INotifyPropertyChanged
    {
        public static readonly DependencyProperty PseudonymProperty =
            DependencyProperty.Register("Pseudonym", typeof(string), typeof(LayerSettingControl), new PropertyMetadata(""));

        public string Pseudonym
        {
            get { return (string)GetValue(PseudonymProperty); }
            set { SetValue(PseudonymProperty, value); }
        }

        public string LocalizedPseudonym
        {
            get
            {
                if (string.IsNullOrEmpty(Pseudonym))
                    return string.Empty;

                var key = $"Layer_{Pseudonym}";
                return LocalizationManager.Instance.GetString(key);
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public LayerSettingControl()
        {
            InitializeComponent();

            // Subscribe to language changes
            LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
        }

        private void OnLanguageChanged(object? sender, EventArgs e)
        {
            // Notify about localized pseudonym change
            OnPropertyChanged(nameof(LocalizedPseudonym));
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            if (((System.Windows.Controls.Button)sender).CommandParameter is LayerSetting layerSetting)
            {
                layerSetting.ResetSettings();
            }
        }

        private void CustomTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = (System.Windows.Controls.TextBox)sender;
            if (textBox.DataContext is not LayerSetting) return;

            var originalText = textBox.Text;
            var cleanedText = LayerNameValidator.CleanAndValidate(originalText);

            if (cleanedText != originalText)
            {
                var caretIndex = textBox.CaretIndex;
                textBox.Text = cleanedText;
                textBox.CaretIndex = Math.Min(caretIndex, cleanedText.Length);
            }
        }
    }
}