using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using MessageBox = System.Windows.MessageBox;

namespace FPECORE;

public partial class OverrideQuantityDialog : Window
{
    public OverrideQuantityDialog()
    {
        InitializeComponent();
        Loaded += OverrideQuantityDialog_Loaded;
    }

    public int? NewQuantity { get; private set; }

    private void OverrideQuantityDialog_Loaded(object sender, RoutedEventArgs e)
    {
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(quantityTextBox.Text, out var quantity) && quantity > 0)
        {
            NewQuantity = quantity;
            DialogResult = true;
        }
        else
        {
            MessageBox.Show("Введите допустимое положительное число.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void QuantityTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !IsTextAllowed(e.Text);
    }

    private static bool IsTextAllowed(string text)
    {
        return Regex.IsMatch(text, "^[0-9]+$");
    }
}