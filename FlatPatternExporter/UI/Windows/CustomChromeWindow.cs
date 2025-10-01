using System.Windows;

namespace FlatPatternExporter.UI.Windows;

public class CustomChromeWindow : Window
{
    public CustomChromeWindow()
    {
        Loaded += OnWindowLoaded;
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        var template = Template;
        if (template != null)
        {
            if (template.FindName("PART_MinimizeButton", this) is System.Windows.Controls.Button minButton)
            {
                minButton.Click += (s, args) => WindowState = WindowState.Minimized;
            }

            if (template.FindName("PART_MaximizeRestoreButton", this) is System.Windows.Controls.Button maxButton)
            {
                maxButton.Click += (s, args) =>
                {
                    WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
                };
            }

            if (template.FindName("PART_CloseButton", this) is System.Windows.Controls.Button closeButton)
            {
                closeButton.Click += (s, args) => Close();
            }
        }
    }
}
