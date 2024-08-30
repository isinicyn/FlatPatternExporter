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
    }
}
