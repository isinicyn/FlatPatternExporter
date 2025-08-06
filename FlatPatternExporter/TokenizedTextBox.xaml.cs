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
    public class TokenElement
    {
        public string Name { get; set; } = string.Empty;
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public Border? VisualElement { get; set; }
        public bool IsCustomText { get; set; } = false;
    }

    public partial class TokenizedTextBox : System.Windows.Controls.UserControl, INotifyPropertyChanged
    {
        private List<TokenElement> _tokenElements = new List<TokenElement>();

        public static readonly DependencyProperty TextTemplateProperty =
            DependencyProperty.Register(nameof(Template), typeof(string), typeof(TokenizedTextBox),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                    OnTemplateChanged));

        public static readonly DependencyProperty AvailableTokensProperty =
            DependencyProperty.Register(nameof(AvailableTokens), typeof(ObservableCollection<string>),
                typeof(TokenizedTextBox),
                new PropertyMetadata(new ObservableCollection<string>(), OnAvailableTokensChanged));

        public static readonly DependencyProperty TokenServiceProperty =
            DependencyProperty.Register(nameof(TokenService), typeof(TokenService), typeof(TokenizedTextBox),
                new PropertyMetadata(null, OnTokenServiceChanged));

        private static readonly Regex TokenRegex = new(@"\{(\w+)\}", RegexOptions.Compiled);
        private static readonly Regex CustomTextRegex = new(@"\{CUSTOM:([^}]+)\}", RegexOptions.Compiled);

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

        public TokenService? TokenService
        {
            get => (TokenService?)GetValue(TokenServiceProperty);
            set => SetValue(TokenServiceProperty, value);
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

        private static void OnTokenServiceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TokenizedTextBox control) control.UpdateContent();
        }

        private void UpdateContent()
        {
            if (TokenContainer == null) return;

            TokenContainer.Children.Clear();
            _tokenElements.Clear();

            if (string.IsNullOrEmpty(Template)) return;

            // Находим все токены и кастомный текст
            var allMatches = new List<(Match match, bool isCustom)>();
            
            // Добавляем обычные токены
            foreach (Match match in TokenRegex.Matches(Template))
                allMatches.Add((match, false));
            
            // Добавляем кастомные токены
            foreach (Match match in CustomTextRegex.Matches(Template))
                allMatches.Add((match, true));
            
            // Сортируем по позиции в строке
            allMatches.Sort((a, b) => a.match.Index.CompareTo(b.match.Index));
            
            for (int i = 0; i < allMatches.Count; i++)
            {
                var (match, isCustom) = allMatches[i];
                var tokenElement = new TokenElement
                {
                    Name = match.Groups[1].Value,
                    StartIndex = match.Index,
                    EndIndex = match.Index + match.Length - 1,
                    IsCustomText = isCustom
                };
                
                AddTokenElement(tokenElement, i);
                _tokenElements.Add(tokenElement);
            }
        }


        private void AddTokenElement(TokenElement tokenElement, int index)
        {
            var border = new Border
            {
                Style = FindResource("TokenBlockStyle") as Style,
                Tag = tokenElement.IsCustomText ? "CustomText" : null
            };

            var textBlock = new TextBlock
            {
                Text = tokenElement.IsCustomText ? $"{tokenElement.Name}" : tokenElement.Name,
                Style = FindResource("TokenTextStyle") as Style
            };

            border.Child = textBlock;
            tokenElement.VisualElement = border;

            border.MouseDown += (s, e) =>
            {
                if (e.RightButton == MouseButtonState.Pressed) RemoveTokenByIndex(index);
            };

            TokenContainer.Children.Add(border);
        }

        private void RemoveTokenByIndex(int index)
        {
            if (index < 0 || index >= _tokenElements.Count) return;
            
            var tokenElement = _tokenElements[index];
            var tokenText = tokenElement.IsCustomText 
                ? $"{{CUSTOM:{tokenElement.Name}}}"
                : $"{{{tokenElement.Name}}}";
            
            var newTemplate = Template.Remove(tokenElement.StartIndex, tokenText.Length);
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

        public void AddCustomText(string customText)
        {
            if (string.IsNullOrEmpty(customText)) return;
            
            // Добавляем кастомный текст как специальный токен
            var customToken = $"{{CUSTOM:{customText}}}";
            Template += customToken;
        }
    }
}