using System.Windows;
using System.Windows.Controls;

namespace FlatPatternExporter
{
    public partial class LayerSettingControl : System.Windows.Controls.UserControl
    {
        public static readonly DependencyProperty PseudonymProperty =
            DependencyProperty.Register("Pseudonym", typeof(string), typeof(LayerSettingControl), new PropertyMetadata(string.Empty));

        public string Pseudonym
        {
            get { return (string)GetValue(PseudonymProperty); }
            set { SetValue(PseudonymProperty, value); }
        }

        private static readonly char[] InvalidCharacters = new[] { '<', '>', '/', '\\', '\"', ':', ';', '?', '*', '|', ',', '=' };
        private bool suppressTextChanged = false;

        public LayerSettingControl()
        {
            InitializeComponent();
        }

        private void ResetLayerSettings_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is LayerSetting layerSetting)
            {
                layerSetting.IsVisible = true;
                layerSetting.CustomName = string.Empty;
                layerSetting.SelectedColor = "White";
                layerSetting.SelectedLineType = "Default";
            }
        }

        private void CustomTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (suppressTextChanged)
            {
                return;
            }

            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox == null) return;

            string originalText = textBox.Text.ToUpper();
            string cleanedText = RemoveInvalidCharacters(originalText);

            if (cleanedText != originalText)
            {
                suppressTextChanged = true; // Подавление дальнейших изменений текста

                int originalCaretIndex = textBox.CaretIndex;
                textBox.Text = cleanedText;
                textBox.CaretIndex = Math.Min(originalCaretIndex - 1, cleanedText.Length);

                suppressTextChanged = false;

                System.Windows.MessageBox.Show(
                    "Недопустимые символы в имени слоя.\nВ именах слоев не допускается употребление следующих символов:\n<>/\\\"\":;?*|,=",
                    "Ошибка ввода",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            else
            {
                // Приведение текста к верхнему регистру
                if (textBox.Text != originalText)
                {
                    suppressTextChanged = true;
                    textBox.Text = originalText;
                    textBox.CaretIndex = textBox.Text.Length;
                    suppressTextChanged = false;
                }
            }
        }

        private string RemoveInvalidCharacters(string input)
        {
            foreach (var c in InvalidCharacters)
            {
                input = input.Replace(c.ToString(), string.Empty);
            }
            return input;
        }
    }
}