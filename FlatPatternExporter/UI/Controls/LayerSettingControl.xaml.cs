using System.Windows;
using System.Windows.Controls;

namespace FlatPatternExporter
{
    public partial class LayerSettingControl : System.Windows.Controls.UserControl
    {
        public static readonly DependencyProperty PseudonymProperty =
            DependencyProperty.Register("Pseudonym", typeof(string), typeof(LayerSettingControl), new PropertyMetadata(""));

        public string Pseudonym
        {
            get { return (string)GetValue(PseudonymProperty); }
            set { SetValue(PseudonymProperty, value); }
        }

        public LayerSettingControl()
        {
            InitializeComponent();
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