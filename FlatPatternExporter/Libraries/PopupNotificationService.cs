using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace WpfToolkit;

/// <summary>
/// Service for displaying temporary popup notifications
/// </summary>
public static class PopupNotificationService
{
    private static readonly ConditionalWeakTable<FrameworkElement, NotificationState> _activeNotifications = new();

    private sealed class NotificationState
    {
        public Popup Popup { get; init; } = null!;
        public DispatcherTimer Timer { get; init; } = null!;

        public void Cleanup()
        {
            Timer.Stop();
            Popup.IsOpen = false;
        }
    }

    /// <summary>
    /// Shows a temporary notification popup with the specified message
    /// </summary>
    /// <param name="element">Element on which the notification will be displayed</param>
    /// <param name="message">Message text</param>
    /// <param name="durationSeconds">Display duration in seconds (default 1.5)</param>
    /// <param name="placement">Popup placement (default Mouse - centered above cursor)</param>
    /// <param name="verticalOffset">Vertical offset in pixels (default -50)</param>
    /// <param name="horizontalOffset">Horizontal offset in pixels (default 0)</param>
    public static void ShowNotification(
        FrameworkElement element,
        string message,
        double durationSeconds = 1.5,
        PlacementMode placement = PlacementMode.Mouse,
        double verticalOffset = -50,
        double horizontalOffset = 0)
    {
        if (element == null || string.IsNullOrEmpty(message))
            return;

        // Close previous notification
        if (_activeNotifications.TryGetValue(element, out var existing))
            existing.Cleanup();

        // Create content with fallback styles
        var textBlock = new TextBlock
        {
            Text = message,
            Style = TryFindResource<Style>("NotificationTextStyle")
        };

        var border = new Border
        {
            Style = TryFindResource<Style>("NotificationContentStyle"),
            Child = textBlock,
            Opacity = 0
        };

        var popup = new Popup
        {
            Child = border,
            PlacementTarget = element,
            Placement = placement,
            AllowsTransparency = true,
            StaysOpen = true
        };

        // Calculate offset on Opened
        popup.Opened += OnPopupOpened;

        void OnPopupOpened(object? s, EventArgs e)
        {
            popup.Opened -= OnPopupOpened;

            border.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
            popup.HorizontalOffset = placement == PlacementMode.Mouse
                ? horizontalOffset - border.DesiredSize.Width / 2
                : horizontalOffset;
            popup.VerticalOffset = verticalOffset;

            // Fade in animation
            border.BeginAnimation(UIElement.OpacityProperty,
                new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200)));
        }

        popup.IsOpen = true;

        // Auto-close timer
        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
        timer.Tick += OnTimerTick;

        void OnTimerTick(object? s, EventArgs e)
        {
            timer.Tick -= OnTimerTick;
            timer.Stop();

            // Fade out animation
            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(150));
            fadeOut.Completed += (_, _) =>
            {
                popup.IsOpen = false;
                _activeNotifications.Remove(element);
            };
            border.BeginAnimation(UIElement.OpacityProperty, fadeOut);
        }

        _activeNotifications.AddOrUpdate(element, new NotificationState { Popup = popup, Timer = timer });
        timer.Start();
    }

    /// <summary>
    /// Closes active notification for the specified element
    /// </summary>
    /// <param name="element">Element for which to close the notification</param>
    public static void CloseNotification(FrameworkElement element)
    {
        if (_activeNotifications.TryGetValue(element, out var state))
        {
            state.Cleanup();
            _activeNotifications.Remove(element);
        }
    }

    private static T? TryFindResource<T>(string key) where T : class
    {
        try
        {
            return System.Windows.Application.Current.TryFindResource(key) as T;
        }
        catch
        {
            return null;
        }
    }
}
