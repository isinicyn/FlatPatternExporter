using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace FlatPatternExporter.UI.Controls;

public partial class CustomTitleBar : System.Windows.Controls.UserControl
{
    private Window? _parentWindow;

    public CustomTitleBar()
    {
        InitializeComponent();
        Loaded += CustomTitleBar_Loaded;
        MouseLeftButtonDown += CustomTitleBar_MouseLeftButtonDown;
    }

    #region Dependency Properties

    public static readonly DependencyProperty TitleProperty =
        DependencyProperty.Register(nameof(Title), typeof(string), typeof(CustomTitleBar),
            new PropertyMetadata(string.Empty));

    public string Title
    {
        get => (string)GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    public static readonly DependencyProperty IconProperty =
        DependencyProperty.Register(nameof(Icon), typeof(ImageSource), typeof(CustomTitleBar),
            new PropertyMetadata(null));

    public ImageSource? Icon
    {
        get => (ImageSource?)GetValue(IconProperty);
        set => SetValue(IconProperty, value);
    }

    public static readonly DependencyProperty ShowMinimizeButtonProperty =
        DependencyProperty.Register(nameof(ShowMinimizeButton), typeof(bool), typeof(CustomTitleBar),
            new PropertyMetadata(true, OnShowMinimizeButtonChanged));

    public bool ShowMinimizeButton
    {
        get => (bool)GetValue(ShowMinimizeButtonProperty);
        set => SetValue(ShowMinimizeButtonProperty, value);
    }

    public static readonly DependencyProperty ShowMaximizeButtonProperty =
        DependencyProperty.Register(nameof(ShowMaximizeButton), typeof(bool), typeof(CustomTitleBar),
            new PropertyMetadata(true, OnShowMaximizeButtonChanged));

    public bool ShowMaximizeButton
    {
        get => (bool)GetValue(ShowMaximizeButtonProperty);
        set => SetValue(ShowMaximizeButtonProperty, value);
    }

    #endregion

    #region Event Handlers

    private void CustomTitleBar_Loaded(object sender, RoutedEventArgs e)
    {
        _parentWindow = Window.GetWindow(this);
        if (_parentWindow != null)
        {
            _parentWindow.StateChanged += ParentWindow_StateChanged;
            UpdateMaximizeRestoreButton();
            UpdateButtonVisibility();
        }
    }

    private void CustomTitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_parentWindow == null) return;

        if (e.ClickCount == 1)
        {
            _parentWindow.DragMove();
        }
        else if (e.ClickCount == 2 && ShowMaximizeButton)
        {
            ToggleMaximizeRestore();
        }
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        if (_parentWindow != null)
        {
            _parentWindow.WindowState = WindowState.Minimized;
        }
    }

    private void MaximizeRestoreButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleMaximizeRestore();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        _parentWindow?.Close();
    }

    private void ParentWindow_StateChanged(object? sender, EventArgs e)
    {
        UpdateMaximizeRestoreButton();
    }

    private static void OnShowMinimizeButtonChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is CustomTitleBar titleBar)
        {
            titleBar.UpdateButtonVisibility();
        }
    }

    private static void OnShowMaximizeButtonChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is CustomTitleBar titleBar)
        {
            titleBar.UpdateButtonVisibility();
        }
    }

    #endregion

    #region Private Methods

    private void ToggleMaximizeRestore()
    {
        if (_parentWindow != null)
        {
            _parentWindow.WindowState = _parentWindow.WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
        }
    }

    private void UpdateMaximizeRestoreButton()
    {
        if (_parentWindow?.WindowState == WindowState.Maximized)
        {
            MaximizeRestorePath.Data = Geometry.Parse("M2,0 L10,0 L10,8 L8,8 M0,2 L0,10 L8,10 L8,8 L2,8 L2,2 Z");
        }
        else
        {
            MaximizeRestorePath.Data = Geometry.Parse("M0,0 L10,0 L10,10 L0,10 Z");
        }
    }

    private void UpdateButtonVisibility()
    {
        MinimizeButton.Visibility = ShowMinimizeButton ? Visibility.Visible : Visibility.Collapsed;
        MaximizeRestoreButton.Visibility = ShowMaximizeButton ? Visibility.Visible : Visibility.Collapsed;
    }

    #endregion
}
