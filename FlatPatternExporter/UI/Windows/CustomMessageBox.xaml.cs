using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using FlatPatternExporter.Services;
using WpfToolkit;

namespace FlatPatternExporter.UI.Windows;

public partial class CustomMessageBox : Window
{
    private MessageBoxResult _result = MessageBoxResult.None;
    private readonly LocalizationManager _localizationManager = LocalizationManager.Instance;
    private string _messageText = string.Empty;

    private CustomMessageBox(string message, string title, MessageBoxButton button, MessageBoxImage icon)
    {
        InitializeComponent();

        _messageText = message;
        MessageTextBlock.Text = message;
        TitleBar.Title = title;
        Title = title;

        SetIcon(icon);
        CreateButtons(button);
    }

    public static MessageBoxResult Show(string messageBoxText)
    {
        return Show(messageBoxText, string.Empty, MessageBoxButton.OK, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(string messageBoxText, string caption)
    {
        return Show(messageBoxText, caption, MessageBoxButton.OK, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(string messageBoxText, string caption, MessageBoxButton button)
    {
        return Show(messageBoxText, caption, button, MessageBoxImage.None);
    }

    public static MessageBoxResult Show(string messageBoxText, string caption, MessageBoxButton button, MessageBoxImage icon)
    {
        return Show(null!, messageBoxText, caption, button, icon);
    }

    public static MessageBoxResult Show(Window? owner, string messageBoxText, string caption, MessageBoxButton button, MessageBoxImage icon)
    {
        var messageBox = new CustomMessageBox(messageBoxText, caption, button, icon);

        // Set owner window to ensure proper modal behavior
        if (owner != null)
        {
            messageBox.Owner = owner;
        }
        else
        {
            // Auto-detect active window if owner not specified
            messageBox.Owner = System.Windows.Application.Current?.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive);
        }

        messageBox.ShowDialog();
        return messageBox._result;
    }

    private void SetIcon(MessageBoxImage icon)
    {
        var (geometry, color) = icon switch
        {
            MessageBoxImage.Information => (FindResource("InfoIconGeometry") as Geometry,
                FindResource("AccentBrush") as SolidColorBrush),
            MessageBoxImage.Warning => (FindResource("WarningIconGeometry") as Geometry,
                FindResource("WarningBrush") as SolidColorBrush),
            MessageBoxImage.Error => (FindResource("ErrorIconGeometry") as Geometry,
                FindResource("ErrorBrush") as SolidColorBrush),
            MessageBoxImage.Question => (FindResource("QuestionIconGeometry") as Geometry,
                FindResource("AccentBrush") as SolidColorBrush),
            _ => (null, null)
        };

        if (geometry != null)
        {
            IconPath.Data = geometry;
            IconPath.Stroke = color;
        }
        else
        {
            IconPath.Visibility = Visibility.Collapsed;
        }
    }

    private void CreateButtons(MessageBoxButton button)
    {
        ButtonPanel.Children.Clear();

        switch (button)
        {
            case MessageBoxButton.OK:
                AddButton(_localizationManager.GetString("Button_OK"), MessageBoxResult.OK, isDefault: true);
                break;

            case MessageBoxButton.OKCancel:
                AddButton(_localizationManager.GetString("Button_OK"), MessageBoxResult.OK, isDefault: true);
                AddButton(_localizationManager.GetString("Button_Cancel"), MessageBoxResult.Cancel, isCancel: true);
                break;

            case MessageBoxButton.YesNo:
                AddButton(_localizationManager.GetString("Button_Yes"), MessageBoxResult.Yes, isDefault: true);
                AddButton(_localizationManager.GetString("Button_No"), MessageBoxResult.No);
                break;

            case MessageBoxButton.YesNoCancel:
                AddButton(_localizationManager.GetString("Button_Yes"), MessageBoxResult.Yes, isDefault: true);
                AddButton(_localizationManager.GetString("Button_No"), MessageBoxResult.No);
                AddButton(_localizationManager.GetString("Button_Cancel"), MessageBoxResult.Cancel, isCancel: true);
                break;
        }
    }

    private void AddButton(string content, MessageBoxResult result, bool isDefault = false, bool isCancel = false)
    {
        var button = new System.Windows.Controls.Button
        {
            Content = content,
            Style = FindResource("FileNameActionButtonStyle") as Style,
            MinWidth = 80,
            Margin = new Thickness(5, 0, 0, 0),
            IsDefault = isDefault,
            IsCancel = isCancel
        };

        button.Click += (s, e) =>
        {
            _result = result;
            DialogResult = true;
            Close();
        };

        ButtonPanel.Children.Add(button);
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            System.Windows.Clipboard.SetText(_messageText);
            var message = _localizationManager.GetString("Text_Copied");
            PopupNotificationService.ShowNotification((FrameworkElement)sender, message);
        }
        catch
        {
            // Игнорируем ошибки при копировании в буфер обмена
        }
    }
}
