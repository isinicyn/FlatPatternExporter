using System.Windows;

namespace FPECORE
{
    public partial class CustomIPropertyDialog : Window
    {
        public string CustomPropertyName { get; private set; }

        public CustomIPropertyDialog()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            CustomPropertyName = propertyNameTextBox.Text;
            DialogResult = true;
            Close();
        }

        private void PropertyNameTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Активируем кнопку OK только если текстовое поле не пустое
            okButton.IsEnabled = !string.IsNullOrWhiteSpace(propertyNameTextBox.Text);
        }
    }
}
