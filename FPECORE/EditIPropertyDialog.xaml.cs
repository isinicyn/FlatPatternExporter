using Inventor;
using System.Windows;

namespace FPECORE
{
    public partial class EditIPropertyDialog : Window
    {
        public string PartNumber { get; private set; }
        public string Description { get; private set; }

        public EditIPropertyDialog(string partNumber, string description)
        {
            InitializeComponent();
            PartNumber = partNumber;
            Description = description;
            partNumberTextBox.Text = PartNumber;
            descriptionTextBox.Text = Description;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PartNumber = partNumberTextBox.Text.Trim();
            Description = descriptionTextBox.Text.Trim();
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
