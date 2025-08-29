using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace FlatPatternExporter
{
    /// <summary>
    /// Конвертирует название цвета в объект SolidColorBrush для отображения в UI
    /// </summary>
    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string colorName)
            {
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorName);
                return new SolidColorBrush(color);
            }
            return System.Windows.Media.Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Конвертирует название типа линии в объект Geometry для визуального отображения
    /// </summary>
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

    /// <summary>
    /// Константы для настроек слоев по умолчанию
    /// </summary>
    public static class LayerDefaults
    {
        public const string DefaultCustomName = "";
        public const string DefaultColor = "White";
        public const string DefaultLineType = "Default";
        public const string OuterProfileLayerName = "OuterProfileLayer";
    }

    /// <summary>
    /// Модель настроек отдельного слоя с поддержкой уведомлений об изменениях
    /// </summary>
    public class LayerSetting : INotifyPropertyChanged
    {
        public string DisplayName { get; set; }
        public string LayerName { get; set; }
        public bool CanBeHidden { get; set; }

        private bool isChecked;
        public bool IsChecked
        {
            get => isChecked;
            set
            {
                isChecked = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsCustomNameEnabled));
                OnPropertyChanged(nameof(IsColorAndLineTypeEnabled));
            }
        }

        public bool IsCustomNameEnabled => IsChecked;
        public bool IsColorAndLineTypeEnabled => IsChecked;
        public bool IsEnabled => DisplayName != LayerDefaults.OuterProfileLayerName;

        private string customName = LayerDefaults.DefaultCustomName;
        public string CustomName
        {
            get => customName;
            set
            {
                customName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanReset));
            }
        }

        private string selectedColor = LayerDefaults.DefaultColor;
        public string SelectedColor
        {
            get => selectedColor;
            set
            {
                selectedColor = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanReset));
            }
        }

        private string selectedLineType = LayerDefaults.DefaultLineType;
        public string SelectedLineType
        {
            get => selectedLineType;
            set
            {
                selectedLineType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanReset));
            }
        }

        public LayerSetting(string displayName, string layerName, bool canBeHidden = true)
        {
            DisplayName = displayName;
            LayerName = layerName;
            CanBeHidden = canBeHidden;
            IsChecked = DisplayName == LayerDefaults.OuterProfileLayerName;
        }

        /// <summary>
        /// Сбрасывает все настройки слоя к значениям по умолчанию
        /// </summary>
        public void ResetSettings()
        {
            CustomName = LayerDefaults.DefaultCustomName;
            SelectedColor = LayerDefaults.DefaultColor;
            SelectedLineType = LayerDefaults.DefaultLineType;
            OnPropertyChanged(nameof(CanReset));
        }

        /// <summary>
        /// Проверяет, отличаются ли текущие настройки от значений по умолчанию
        /// </summary>
        public bool HasChanges() =>
            CustomName != LayerDefaults.DefaultCustomName ||
            SelectedColor != LayerDefaults.DefaultColor ||
            SelectedLineType != LayerDefaults.DefaultLineType;

        public bool CanReset => HasChanges();

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Статический помощник для работы с цветами, типами линий и настройками слоев
    /// </summary>
    public static class LayerSettingsHelper
    {
        private static readonly Dictionary<string, string> ColorValues = new()
        {
            { "White", "255;255;255" },
            { "Red", "255;0;0" },
            { "Orange", "255;81;0" },
            { "Yellow", "234;255;0" },
            { "Green", "55;255;0" },
            { "Cyan", "0;195;255" },
            { "Blue", "13;0;255" },
            { "Purple", "255;10;169" },
            { "DarkGray", "169;169;169" },
            { "LightGray", "211;211;211" }
        };

        private static readonly Dictionary<string, string> LineTypeValues = new()
        {
            { "Chain", "37644" },
            { "Continuous", "37633" },
            { "Dash dotted", "37638" },
            { "Dashed double dotted", "37645" },
            { "Dashed hidden", "37641" },
            { "Dashed", "37634" },
            { "Dashed triple dotted", "37647" },
            { "Default", "37648" },
            { "Dotted", "37636" },
            { "Double dash double dotted", "37639" },
            { "Double dashed", "37637" },
            { "Double dashed dotted", "37646" },
            { "Double dash triple dotted", "37640" },
            { "Long dash dotted", "37642" },
            { "Long dashed double dotted", "37635" },
            { "Long dash triple dotted", "37643" }
        };

        private static readonly List<LayerDefinition> LayerDefinitions = new()
        {
            new("BendUpLayer", "IV_BEND"),
            new("BendDownLayer", "IV_BEND_DOWN"),
            new("ToolCenterUpLayer", "IV_TOOL_CENTER"),
            new("ToolCenterDownLayer", "IV_TOOL_CENTER_DOWN"),
            new("ArcCentersLayer", "IV_ARC_CENTERS"),
            new("OuterProfileLayer", "IV_OUTER_PROFILE", CanBeHidden: false),
            new("InteriorProfilesLayer", "IV_INTERIOR_PROFILES"),
            new("FeatureProfilesUpLayer", "IV_FEATURE_PROFILES"),
            new("FeatureProfilesDownLayer", "IV_FEATURE_PROFILES_DOWN"),
            new("AltRepFrontLayer", "IV_ALTREP_FRONT"),
            new("AltRepBackLayer", "IV_ALTREP_BACK"),
            new("UnconsumedSketchesLayer", "IV_UNCONSUMED_SKETCHES"),
            new("TangentLayer", "IV_TANGENT"),
            new("TangentRollLinesLayer", "IV_ROLL_TANGENT"),
            new("RollLinesLayer", "IV_ROLL"),
            new("UnconsumedSketchConstructionLayer", "IV_UNCONSUMED_SKETCH_CONSTRUCTION")
        };

        /// <summary>
        /// Конвертирует название цвета в RGB значение
        /// </summary>
        public static string GetColorValue(string colorName) =>
            ColorValues.TryGetValue(colorName, out string? value) ? value : ColorValues["White"];

        /// <summary>
        /// Конвертирует название типа линии в его соответствующее значение
        /// </summary>
        public static string GetLineTypeValue(string lineTypeName) =>
            LineTypeValues.TryGetValue(lineTypeName, out string? value) ? value : LineTypeValues["Default"];

        /// <summary>
        /// Инициализирует коллекцию настроек слоев с предустановленными значениями
        /// </summary>
        public static ObservableCollection<LayerSetting> InitializeLayerSettings()
        {
            var layerSettings = new ObservableCollection<LayerSetting>();
            
            foreach (var definition in LayerDefinitions)
            {
                layerSettings.Add(new LayerSetting(definition.DisplayName, definition.LayerName, definition.CanBeHidden));
            }
            
            return layerSettings;
        }

        /// <summary>
        /// Возвращает коллекцию доступных цветов
        /// </summary>
        public static ObservableCollection<string> GetAvailableColors() =>
            new(ColorValues.Keys);

        /// <summary>
        /// Возвращает коллекцию доступных типов линий
        /// </summary>
        public static ObservableCollection<string> GetLineTypes() =>
            new(LineTypeValues.Keys);

        /// <summary>
        /// Внутренний класс для определения слоев
        /// </summary>
        private readonly record struct LayerDefinition(string DisplayName, string LayerName, bool CanBeHidden = true);
    }

    /// <summary>
    /// Валидатор имен слоев с автоматической очисткой недопустимых символов
    /// </summary>
    public static class LayerNameValidator
    {
        private static readonly HashSet<char> InvalidCharacters = new() 
        { 
            '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=' 
        };

        private const string ValidationMessage = 
            "Недопустимые символы в имени слоя.\nВ именах слоев не допускается употребление следующих символов:\n<>/\\\":;?*|,=";

        /// <summary>
        /// Очищает строку от недопустимых символов и показывает предупреждение при необходимости
        /// </summary>
        public static string CleanAndValidate(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var cleaned = new StringBuilder(input.Length);
            bool hasInvalidChars = false;

            foreach (char c in input)
            {
                if (InvalidCharacters.Contains(c))
                {
                    hasInvalidChars = true;
                }
                else
                {
                    cleaned.Append(c);
                }
            }

            if (hasInvalidChars)
            {
                ShowValidationError();
            }

            return cleaned.ToString();
        }

        /// <summary>
        /// Отображает диалог с информацией об ошибке валидации
        /// </summary>
        private static void ShowValidationError()
        {
            System.Windows.MessageBox.Show(
                ValidationMessage,
                "Ошибка ввода",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }
}