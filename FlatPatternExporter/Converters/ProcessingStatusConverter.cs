using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Converters;

public class ProcessingStatusConverter : IMultiValueConverter
{
    public static readonly ProcessingStatusConverter Instance = new();

    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length < 1 || values[0] is not ProcessingStatus status)
            return string.Empty;

        return status switch
        {
            ProcessingStatus.NotProcessed => LocalizationManager.Instance.GetString("ProcessingStatus_NotProcessed"),
            ProcessingStatus.Pending => LocalizationManager.Instance.GetString("ProcessingStatus_Pending"),
            ProcessingStatus.Success => LocalizationManager.Instance.GetString("ProcessingStatus_Success"),
            ProcessingStatus.Skipped => LocalizationManager.Instance.GetString("ProcessingStatus_Skipped"),
            _ => string.Empty
        };
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException("ProcessingStatusConverter does not support ConvertBack");
    }
}
