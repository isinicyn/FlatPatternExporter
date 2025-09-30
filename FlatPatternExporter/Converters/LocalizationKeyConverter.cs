using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Converters;

[ValueConversion(typeof(CultureInfo), typeof(string))]
public class LocalizationKeyConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (parameter is string key && !string.IsNullOrEmpty(key))
        {
            return LocalizationManager.Instance.GetString(key);
        }

        return parameter?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException("LocalizationKeyConverter does not support ConvertBack");
    }
}