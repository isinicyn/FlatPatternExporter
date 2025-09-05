using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace FlatPatternExporter.Converters;

public class LineTypeToGeometryConverter : IValueConverter
{
    private static readonly Dictionary<string, string> LineTypeGeometry = new()
    {
        { "Continuous", "M0,8 L50,8" },
        { "Chain", "M0,8 L15,8 M20,8 L30,8 M35,8 L50,8" },
        { "Dashed", "M0,8 L12,8 M18,8 L30,8 M36,8 L48,8" },
        { "Dotted", "M0,8 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2" },
        { "Dash dotted", "M0,8 L12,8 M18,8 h2 M24,8 L36,8 M42,8 h2" },
        { "Dashed double dotted", "M0,8 L10,8 M14,8 h2 m2,0 h2 M22,8 L32,8 M36,8 h2 m2,0 h2" },
        { "Dashed hidden", "M0,8 L15,8 M35,8 L50,8" },
        { "Dashed triple dotted", "M0,8 L8,8 M12,8 h2 m2,0 h2 m2,0 h2 M28,8 L36,8 M40,8 h2 m2,0 h2" },
        { "Default", "M0,8 L50,8" },
        { "Double dash double dotted", "M0,8 L8,8 M12,8 h2 m3,0 h2 M24,8 L32,8 M36,8 h2 m3,0 h2" },
        { "Double dashed", "M0,8 L10,8 M15,8 L25,8 M30,8 L40,8 M45,8 L50,8" },
        { "Double dashed dotted", "M0,8 L10,8 M14,8 h2 M18,8 L28,8 M32,8 h2 M36,8 L46,8" },
        { "Double dash triple dotted", "M0,8 L10,8 M14,8 h2 m2,0 h2 m2,0 h2 M30,8 L40,8 M44,8 h2 m2,0 h2" },
        { "Long dash dotted", "M0,8 L20,8 M25,8 h2 M30,8 L50,8" },
        { "Long dashed double dotted", "M0,8 L18,8 M22,8 h2 m2,0 h2 M32,8 L50,8" },
        { "Long dash triple dotted", "M0,8 L18,8 M22,8 h2 m2,0 h2 m2,0 h2 M38,8 L50,8" }
    };

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string lineTypeName)
        {
            return LineTypeGeometry.TryGetValue(lineTypeName, out string? geometryPath) 
                ? Geometry.Parse(geometryPath) 
                : Geometry.Parse(LineTypeGeometry["Default"]);
        }
        return Geometry.Parse(LineTypeGeometry["Default"]);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}