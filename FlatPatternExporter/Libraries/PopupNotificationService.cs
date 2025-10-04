using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;

namespace WpfToolkit;

/// <summary>
/// Service for displaying temporary popup notifications
/// </summary>
public static class PopupNotificationService
{
    private static readonly Dictionary<FrameworkElement, (Popup popup, DispatcherTimer timer)> _activeNotifications = new();

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

        // Close previous notification for this element if exists
        if (_activeNotifications.TryGetValue(element, out var existing))
        {
            existing.timer.Stop();
            existing.popup.IsOpen = false;
            _activeNotifications.Remove(element);
        }

        // Create notification content using XAML styles
        var textBlock = new TextBlock
        {
            Text = message,
            Style = (Style)System.Windows.Application.Current.FindResource("NotificationTextStyle")
        };

        var border = new Border
        {
            Style = (Style)System.Windows.Application.Current.FindResource("NotificationContentStyle"),
            Child = textBlock
        };

        // Create popup
        var popup = new Popup
        {
            Child = border,
            PlacementTarget = element,
            Placement = placement,
            AllowsTransparency = true,
            StaysOpen = true
        };

        // Center popup horizontally above cursor when using Mouse placement
        popup.Opened += (s, e) =>
        {
            border.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
            if (placement == PlacementMode.Mouse)
            {
                popup.HorizontalOffset = horizontalOffset - (border.DesiredSize.Width / 2);
                popup.VerticalOffset = verticalOffset;
            }
            else
            {
                popup.HorizontalOffset = horizontalOffset;
                popup.VerticalOffset = verticalOffset;
            }
        };

        // Show popup
        popup.IsOpen = true;

        // Create timer to close popup
        var timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(durationSeconds)
        };

        timer.Tick += (s, args) =>
        {
            timer.Stop();
            popup.IsOpen = false;
            _activeNotifications.Remove(element);
        };

        _activeNotifications[element] = (popup, timer);
        timer.Start();
    }
}
