using System.Text.RegularExpressions;
using System.Windows;

namespace FPECORE
{
    public partial class OverrideQuantityDialog : Window
    {
        public int? NewQuantity { get; private set; }

        public OverrideQuantityDialog()
        {
            InitializeComponent();
            this.Loaded += OverrideQuantityDialog_Loaded;
        }

        private void OverrideQuantityDialog_Loaded(object sender, RoutedEventArgs e)
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(quantityTextBox.Text, out int quantity) && quantity > 0)
            {
                NewQuantity = quantity;
                this.DialogResult = true;
            }
            else
            {
                System.Windows.MessageBox.Show("Введите допустимое положительное число.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private void QuantityTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            return Regex.IsMatch(text, "^[0-9]+$");
        }
    }
}
