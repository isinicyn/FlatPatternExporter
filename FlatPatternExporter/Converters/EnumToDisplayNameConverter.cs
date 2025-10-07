using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Converters;

public class EnumToDisplayNameConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length < 1 || values[0] == null)
            return string.Empty;

        return values[0] switch
        {
            SplineReplacementType splineType => SplineReplacementMapping.GetDisplayName(splineType),
            CsvDelimiterType csvDelimiter => CsvDelimiterMapping.GetDisplayName(csvDelimiter),
            _ => values[0].ToString() ?? string.Empty
        };
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}