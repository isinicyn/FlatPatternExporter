using System.Windows;

namespace FlatPatternExporter
{
    public partial class TextWithFxIndicator : System.Windows.Controls.UserControl
    {
        public TextWithFxIndicator()
        {
            InitializeComponent();
        }

        public static readonly DependencyProperty TextProperty = DependencyProperty.Register(
            nameof(Text), typeof(string), typeof(TextWithFxIndicator), new PropertyMetadata(string.Empty));

        public string Text
        {
            get => (string)GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }

        public static readonly DependencyProperty IsExpressionProperty = DependencyProperty.Register(
            nameof(IsExpression), typeof(bool), typeof(TextWithFxIndicator), new PropertyMetadata(false));

        public bool IsExpression
        {
            get => (bool)GetValue(IsExpressionProperty);
            set => SetValue(IsExpressionProperty, value);
        }
    }
}
