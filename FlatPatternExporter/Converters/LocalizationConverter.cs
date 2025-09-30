using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Converters;

[ValueConversion(typeof(string), typeof(string))]
public class LocalizationConverter : IValueConverter
{
    public static readonly LocalizationConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string key && !string.IsNullOrEmpty(key))
        {
            return LocalizationManager.Instance.GetString(key);
        }

        return value?.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException("LocalizationConverter does not support ConvertBack");
    }
}