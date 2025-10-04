using System.Globalization;
using System.Windows.Data;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Converters;

[ValueConversion(typeof(ProcessingStatus), typeof(string))]
public class ProcessingStatusConverter : IValueConverter
{
    public static readonly ProcessingStatusConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is ProcessingStatus status)
        {
            return status switch
            {
                ProcessingStatus.NotProcessed => LocalizationManager.Instance.GetString("ProcessingStatus_NotProcessed"),
                ProcessingStatus.Pending => LocalizationManager.Instance.GetString("ProcessingStatus_Pending"),
                ProcessingStatus.Success => LocalizationManager.Instance.GetString("ProcessingStatus_Success"),
                ProcessingStatus.Skipped => LocalizationManager.Instance.GetString("ProcessingStatus_Skipped"),
                _ => string.Empty
            };
        }

        return string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException("ProcessingStatusConverter does not support ConvertBack");
    }
}
