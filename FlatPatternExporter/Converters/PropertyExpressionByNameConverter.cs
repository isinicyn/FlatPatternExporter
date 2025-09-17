using System.Globalization;
using System.Windows;
using System.Windows.Data;
using FlatPatternExporter.UI.Windows;

namespace FlatPatternExporter.Converters;

public class PropertyExpressionByNameConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values is [PartData partData, _, string propPath])
        {
            return partData.IsPropertyExpression(propPath) ? Visibility.Visible : Visibility.Collapsed;
        }
        return Visibility.Collapsed;
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        return Array.Empty<object>();
    }
}