using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace FlatPatternExporter
{
    public partial class TokenizedTextBox : System.Windows.Controls.UserControl, INotifyPropertyChanged
    {
        public static readonly DependencyProperty TextTemplateProperty =
            DependencyProperty.Register(nameof(Template), typeof(string), typeof(TokenizedTextBox),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                    OnTemplateChanged));

        public static readonly DependencyProperty AvailableTokensProperty =
            DependencyProperty.Register(nameof(AvailableTokens), typeof(ObservableCollection<string>),
                typeof(TokenizedTextBox),
                new PropertyMetadata(new ObservableCollection<string>(), OnAvailableTokensChanged));

        private string _template = string.Empty;

        public TokenizedTextBox()
        {
            InitializeComponent();
        }

        public new string Template
        {
            get => (string)GetValue(TextTemplateProperty);
            set => SetValue(TextTemplateProperty, value);
        }

        public ObservableCollection<string> AvailableTokens
        {
            get => (ObservableCollection<string>)GetValue(AvailableTokensProperty);
            set => SetValue(AvailableTokensProperty, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static void OnTemplateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TokenizedTextBox control) control.UpdateContent();
        }

        private static void OnAvailableTokensChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TokenizedTextBox control) control.UpdateContent();
        }

        private void UpdateContent()
        {
            if (TokenContainer == null) return;

            TokenContainer.Children.Clear();

            if (string.IsNullOrEmpty(Template)) return;

            var pattern = @"\{(\w+)\}";
            var matches = Regex.Matches(Template, pattern);
            
            foreach (Match match in matches)
            {
                var tokenName = match.Groups[1].Value;
                AddTokenElement(tokenName);
            }
        }


        private void AddTokenElement(string tokenName)
        {
            var border = new Border
            {
                Style = FindResource("TokenBlockStyle") as Style
            };

            var textBlock = new TextBlock
            {
                Text = tokenName,
                Style = FindResource("TokenTextStyle") as Style
            };

            border.Child = textBlock;

            border.MouseDown += (s, e) =>
            {
                if (e.RightButton == MouseButtonState.Pressed) RemoveToken(tokenName);
            };

            TokenContainer.Children.Add(border);
        }

        private void RemoveToken(string tokenName)
        {
            var pattern = @"\{" + Regex.Escape(tokenName) + @"\}";
            var newTemplate = Regex.Replace(Template, pattern, "", RegexOptions.IgnoreCase);
            Template = newTemplate;
        }

        public void AddToken(string tokenName)
        {
            if (string.IsNullOrEmpty(tokenName)) return;
            
            Template += tokenName; // Токены уже приходят в формате {TokenName}
        }

        private void TokenContainer_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.StringFormat))
            {
                var token = e.Data.GetData(System.Windows.DataFormats.StringFormat) as string;
                if (!string.IsNullOrEmpty(token)) AddToken(token);
            }
        }

        private void TokenContainer_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(System.Windows.DataFormats.StringFormat) ? System.Windows.DragDropEffects.Copy : System.Windows.DragDropEffects.None;
            e.Handled = true;
        }
    }
}