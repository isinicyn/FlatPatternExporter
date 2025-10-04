using System.Globalization;
using System.Windows;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;

namespace FlatPatternExporter;

/// <summary>
///     Interaction logic for App.xaml
/// </summary>
public partial class App
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Load settings ONCE at application startup
        var settings = SettingsManager.LoadSettings();

        // Apply global settings (theme and language)
        ApplyTheme(settings.Interface.SelectedTheme);
        ApplyLanguage(settings.Interface.SelectedLanguage);

        // Create and show main window with settings
        var mainWindow = new FlatPatternExporterMainWindow();
        mainWindow.ApplySettings(settings);
        mainWindow.Show();
    }

    private void ApplyTheme(string theme)
    {
        try
        {
            var themeFileName = theme == "Light"
                ? "ColorResources.xaml"
                : "DarkTheme.xaml";
            var themeUri = new Uri($"Styles/{themeFileName}", UriKind.Relative);

            var mergedDictionaries = Current.Resources.MergedDictionaries;
            var existingTheme = mergedDictionaries.FirstOrDefault(d =>
                d.Source?.OriginalString.Contains("ColorResources.xaml") == true ||
                d.Source?.OriginalString.Contains("DarkTheme.xaml") == true);

            if (existingTheme != null)
            {
                var index = mergedDictionaries.IndexOf(existingTheme);
                mergedDictionaries.RemoveAt(index);
                mergedDictionaries.Insert(index, new ResourceDictionary { Source = themeUri });
            }
        }
        catch
        {
            // If theme application fails, use default theme from App.xaml
        }
    }

    private void ApplyLanguage(string languageCode)
    {
        try
        {
            var savedLanguage = SupportedLanguages.All
                .FirstOrDefault(lang => lang.Code == languageCode);

            if (savedLanguage != null)
            {
                LocalizationManager.Instance.CurrentCulture = savedLanguage.Culture;
            }
        }
        catch
        {
            // If language application fails, use default language
        }
    }
}