using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace FlatPatternExporter.UI.Controls;

public enum UpdateButtonState
{
    None,
    UpdateAvailable,
    UpToDate,
    Error
}

public partial class CustomTitleBar : System.Windows.Controls.UserControl
{
    private Window? _parentWindow;
    private bool _isMouseDown;
    private System.Windows.Point _startPosition;

    public CustomTitleBar()
    {
        InitializeComponent();
        Loaded += CustomTitleBar_Loaded;
        MouseLeftButtonDown += CustomTitleBar_MouseLeftButtonDown;
        MouseMove += CustomTitleBar_MouseMove;
        MouseLeftButtonUp += CustomTitleBar_MouseLeftButtonUp;
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

    public static readonly DependencyProperty ContentAreaProperty =
        DependencyProperty.Register(nameof(ContentArea), typeof(object), typeof(CustomTitleBar),
            new PropertyMetadata(null));

    public object? ContentArea
    {
        get => GetValue(ContentAreaProperty);
        set => SetValue(ContentAreaProperty, value);
    }

    public static readonly DependencyProperty TitleAlignmentProperty =
        DependencyProperty.Register(nameof(TitleAlignment), typeof(System.Windows.HorizontalAlignment), typeof(CustomTitleBar),
            new PropertyMetadata(System.Windows.HorizontalAlignment.Left));

    public System.Windows.HorizontalAlignment TitleAlignment
    {
        get => (System.Windows.HorizontalAlignment)GetValue(TitleAlignmentProperty);
        set => SetValue(TitleAlignmentProperty, value);
    }

    public static readonly DependencyProperty ShowUpdateButtonProperty =
        DependencyProperty.Register(nameof(ShowUpdateButton), typeof(bool), typeof(CustomTitleBar),
            new PropertyMetadata(false, OnShowUpdateButtonChanged));

    public bool ShowUpdateButton
    {
        get => (bool)GetValue(ShowUpdateButtonProperty);
        set => SetValue(ShowUpdateButtonProperty, value);
    }

    public static readonly DependencyProperty UpdateTooltipProperty =
        DependencyProperty.Register(nameof(UpdateTooltip), typeof(string), typeof(CustomTitleBar),
            new PropertyMetadata(string.Empty));

    public string UpdateTooltip
    {
        get => (string)GetValue(UpdateTooltipProperty);
        set => SetValue(UpdateTooltipProperty, value);
    }

    public static readonly DependencyProperty UpdateStateProperty =
        DependencyProperty.Register(nameof(UpdateState), typeof(UpdateButtonState), typeof(CustomTitleBar),
            new PropertyMetadata(UpdateButtonState.None));

    public UpdateButtonState UpdateState
    {
        get => (UpdateButtonState)GetValue(UpdateStateProperty);
        set => SetValue(UpdateStateProperty, value);
    }

    public static readonly DependencyProperty ReserveUpdateButtonSpaceProperty =
        DependencyProperty.Register(nameof(ReserveUpdateButtonSpace), typeof(bool), typeof(CustomTitleBar),
            new PropertyMetadata(false));

    public bool ReserveUpdateButtonSpace
    {
        get => (bool)GetValue(ReserveUpdateButtonSpaceProperty);
        set => SetValue(ReserveUpdateButtonSpaceProperty, value);
    }

    #endregion

    #region Events

    public event EventHandler? UpdateButtonClick;

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

        if (e.ClickCount == 2 && ShowMaximizeButton)
        {
            ToggleMaximizeRestore();
            _isMouseDown = false;
        }
        else if (e.ClickCount == 1)
        {
            _isMouseDown = true;
            _startPosition = e.GetPosition(this);
        }
    }

    private void CustomTitleBar_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isMouseDown || _parentWindow == null || e.LeftButton != MouseButtonState.Pressed)
            return;

        var currentPosition = e.GetPosition(this);
        var diff = currentPosition - _startPosition;

        if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
            Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
        {
            _isMouseDown = false;

            if (_parentWindow.WindowState == WindowState.Maximized)
            {
                var relativePosition = _startPosition.X / ActualWidth;

                _parentWindow.WindowState = WindowState.Normal;

                _parentWindow.Left = currentPosition.X - (_parentWindow.RestoreBounds.Width * relativePosition);
                _parentWindow.Top = currentPosition.Y - (_startPosition.Y * 0.5);
            }

            _parentWindow.DragMove();
        }
    }

    private void CustomTitleBar_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        _isMouseDown = false;
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

    private static void OnShowUpdateButtonChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is CustomTitleBar titleBar)
        {
            titleBar.UpdateButtonVisibility();
        }
    }

    private void UpdateButton_Click(object sender, RoutedEventArgs e)
    {
        UpdateButtonClick?.Invoke(this, EventArgs.Empty);
    }

    #endregion

    #region Public Methods

    public void ShowUpdateNotification(string message, double durationSeconds = 2.0)
    {
        if (UpdateButton.Visibility == Visibility.Visible)
        {
            WpfToolkit.PopupNotificationService.ShowNotification(
                UpdateButton,
                message,
                durationSeconds,
                System.Windows.Controls.Primitives.PlacementMode.Bottom,
                verticalOffset: 5,
                horizontalOffset: 0
            );
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
        UpdateButton.Visibility = ShowUpdateButton ? Visibility.Visible : Visibility.Collapsed;
    }

    #endregion
}
