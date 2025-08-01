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

        public LayerSettingControl()
        {
            InitializeComponent();
        }

    }
}