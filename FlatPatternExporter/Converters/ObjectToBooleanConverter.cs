using System.Globalization;
using System.Windows.Data;

namespace FlatPatternExporter.Converters;

public class ObjectToBooleanConverter : IValueConverter
{
    public static readonly ObjectToBooleanConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value != null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return System.Windows.Data.Binding.DoNothing;
    }
}