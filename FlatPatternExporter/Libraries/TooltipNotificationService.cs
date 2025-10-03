using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using ToolTip = System.Windows.Controls.ToolTip;

namespace WpfToolkit;

/// <summary>
/// Service for displaying temporary tooltip notifications
/// </summary>
public static class TooltipNotificationService
{
    private static readonly Dictionary<FrameworkElement, DispatcherTimer> _activeTimers = new();

    /// <summary>
    /// Shows a temporary ToolTip with the specified message
    /// </summary>
    /// <param name="element">Element on which the ToolTip will be displayed</param>
    /// <param name="message">Message text</param>
    /// <param name="durationSeconds">Display duration in seconds (default 1.5)</param>
    /// <param name="placement">ToolTip placement (default Bottom)</param>
    public static void ShowTemporaryTooltip(
        FrameworkElement element,
        string message,
        double durationSeconds = 1.5,
        PlacementMode placement = PlacementMode.Bottom)
    {
        if (element == null || string.IsNullOrEmpty(message))
            return;

        // Stop previous timer for this element if exists
        if (_activeTimers.TryGetValue(element, out var existingTimer))
        {
            existingTimer.Stop();
            if (element.ToolTip is ToolTip existingTooltip)
            {
                existingTooltip.IsOpen = false;
            }
        }

        // Save original ToolTip
        var originalTooltip = element.ToolTip;

        // Create and show temporary ToolTip
        var tooltip = new ToolTip
        {
            Content = message,
            IsOpen = true,
            PlacementTarget = element,
            Placement = placement
        };
        element.ToolTip = tooltip;

        // Create timer to restore original ToolTip
        var timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(durationSeconds)
        };

        timer.Tick += (s, args) =>
        {
            timer.Stop();
            _activeTimers.Remove(element);

            if (element.ToolTip is ToolTip tt)
            {
                tt.IsOpen = false;
            }
            element.ToolTip = originalTooltip;
        };

        _activeTimers[element] = timer;
        timer.Start();
    }
}
