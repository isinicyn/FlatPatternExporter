using System.Globalization;
using System.Windows.Data;

namespace FlatPatternExporter.Converters;

public class EnumToBooleanConverter : IValueConverter
{
    public static readonly EnumToBooleanConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null)
            return false;

        var valueType = value.GetType();
        if (valueType.IsEnum && Enum.TryParse(valueType, parameter.ToString(), out var parameterValue))
        {
            return value.Equals(parameterValue);
        }

        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool boolValue && boolValue && parameter != null && targetType.IsEnum)
        {
            if (Enum.TryParse(targetType, parameter.ToString(), out var result))
            {
                return result;
            }
        }

        return System.Windows.Data.Binding.DoNothing;
    }
}