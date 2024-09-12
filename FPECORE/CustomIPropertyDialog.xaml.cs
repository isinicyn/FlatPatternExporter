using System.Windows;
using System.Windows.Controls;

namespace FPECORE;

public partial class CustomIPropertyDialog : Window
{
    public CustomIPropertyDialog()
    {
        InitializeComponent();
    }

    public string CustomPropertyName { get; private set; }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        CustomPropertyName = propertyNameTextBox.Text;
        DialogResult = true;
        Close();
    }

    private void PropertyNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        // Активируем кнопку OK только если текстовое поле не пустое
        okButton.IsEnabled = !string.IsNullOrWhiteSpace(propertyNameTextBox.Text);
    }
}