using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Converters;

public class DynamicPropertyValueConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values is [PartData partData, string propPath])
        {
            if (propPath.StartsWith("UserDefinedProperties["))
            {
                var key = propPath["UserDefinedProperties[".Length..].TrimEnd(']');
                return partData.UserDefinedProperties.TryGetValue(key, out var v) ? v : string.Empty;
            }
            var pi = typeof(PartData).GetProperty(propPath);
            return pi?.GetValue(partData)?.ToString() ?? string.Empty;
        }
        return string.Empty;
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        return [];
    }
}
