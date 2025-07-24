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

    public class LineTypeToGeometryConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string lineTypeName)
            {
                return lineTypeName switch
                {
                    "Continuous" => Geometry.Parse("M0,8 L50,8"),
                    "Chain" => Geometry.Parse("M0,8 L15,8 M20,8 L30,8 M35,8 L50,8"),
                    "Dashed" => Geometry.Parse("M0,8 L12,8 M18,8 L30,8 M36,8 L48,8"),
                    "Dotted" => Geometry.Parse("M0,8 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2 m4,0 h2"),
                    "Dash dotted" => Geometry.Parse("M0,8 L12,8 M18,8 h2 M24,8 L36,8 M42,8 h2"),
                    "Dashed double dotted" => Geometry.Parse("M0,8 L10,8 M14,8 h2 m2,0 h2 M22,8 L32,8 M36,8 h2 m2,0 h2"),
                    "Dashed hidden" => Geometry.Parse("M0,8 L15,8 M35,8 L50,8"),
                    "Dashed triple dotted" => Geometry.Parse("M0,8 L8,8 M12,8 h2 m2,0 h2 m2,0 h2 M28,8 L36,8 M40,8 h2 m2,0 h2"),
                    "Default" => Geometry.Parse("M0,8 L50,8"),
                    "Double dash double dotted" => Geometry.Parse("M0,8 L8,8 M12,8 h2 m3,0 h2 M24,8 L32,8 M36,8 h2 m3,0 h2"),
                    "Double dashed" => Geometry.Parse("M0,8 L10,8 M15,8 L25,8 M30,8 L40,8 M45,8 L50,8"),
                    "Double dashed dotted" => Geometry.Parse("M0,8 L10,8 M14,8 h2 M18,8 L28,8 M32,8 h2 M36,8 L46,8"),
                    "Double dash triple dotted" => Geometry.Parse("M0,8 L10,8 M14,8 h2 m2,0 h2 m2,0 h2 M30,8 L40,8 M44,8 h2 m2,0 h2"),
                    "Long dash dotted" => Geometry.Parse("M0,8 L20,8 M25,8 h2 M30,8 L50,8"),
                    "Long dashed double dotted" => Geometry.Parse("M0,8 L18,8 M22,8 h2 m2,0 h2 M32,8 L50,8"),
                    "Long dash triple dotted" => Geometry.Parse("M0,8 L18,8 M22,8 h2 m2,0 h2 m2,0 h2 M38,8 L50,8"),
                    _ => Geometry.Parse("M0,8 L50,8"), // Default, в случае ошибки
                };
            }
            return Geometry.Parse("M0,8 L50,8");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class LayerSetting : INotifyPropertyChanged
    {
        public string DisplayName { get; set; } // Название настройки, например, BendUpLayer
        public string LayerName { get; set; } // Стандартное имя слоя, например, IV_BEND
        public bool HasVisibilityOption { get; set; }

        private bool isVisible;
        public bool IsVisible
        {
            get => isVisible;
            set
            {
                isVisible = value;
                OnPropertyChanged();
                // Обновление доступности других опций при изменении видимости
                OnPropertyChanged(nameof(IsCustomNameEnabled));
                OnPropertyChanged(nameof(IsColorAndLineTypeEnabled));
            }
        }

        public bool IsCustomNameEnabled => IsVisible;
        public bool IsColorAndLineTypeEnabled => IsVisible;

        private string customName = string.Empty;
        public string CustomName
        {
            get => customName;
            set
            {
                customName = value;
                OnPropertyChanged();
            }
        }

        private string selectedColor = "White";
        public string SelectedColor
        {
            get => selectedColor;
            set
            {
                selectedColor = value;
                OnPropertyChanged();
            }
        }

        private string selectedLineType = "Default";
        public string SelectedLineType
        {
            get => selectedLineType;
            set
            {
                selectedLineType = value;
                OnPropertyChanged();
            }
        }


        public string OriginalName { get; set; } // Оригинальное имя слоя, например, IV_BEND

        public LayerSetting(string displayName, string layerName, bool hasVisibilityOption = true)
        {
            DisplayName = displayName;
            LayerName = layerName;
            OriginalName = layerName; // Инициализируем OriginalName
            HasVisibilityOption = hasVisibilityOption;
            IsVisible = true; // Видимость по умолчанию
            CustomName = string.Empty;
            SelectedColor = "White"; // Цвет по умолчанию
            SelectedLineType = "Default"; // Тип линии по умолчанию
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public bool IsCheckBoxEnabled => DisplayName != "OuterProfileLayer";
    }

    public class CustomNameForegroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (string.IsNullOrWhiteSpace(value?.ToString()))
            {
                return new SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 160, 160, 160)); // Полупрозрачный цвет для пустого текста
            }
            return new SolidColorBrush(Colors.Black); // Черный текст для пользовательского ввода
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TextEmptyToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Если поле пустое, показываем оригинальное имя слоя
            if (string.IsNullOrWhiteSpace(value?.ToString()))
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public static class LayerSettingsHelper
    {
        // Метод для конвертации имени цвета в RGB значение
        public static string GetColorValue(string colorName)
        {
            var colorDictionary = new Dictionary<string, string>
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

            return colorDictionary.ContainsKey(colorName) ? colorDictionary[colorName] : "255;255;255";
        }

        // Метод для конвертации имени типа линии в его соответствующее значение
        public static string GetLineTypeValue(string lineTypeName)
        {
            var lineTypeDictionary = new Dictionary<string, string>
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

            return lineTypeDictionary.ContainsKey(lineTypeName) ? lineTypeDictionary[lineTypeName] : "37648";
        }

        // Метод для создания строки экспорта параметров слоев
        public static string ExportLayerOptions(ObservableCollection<LayerSetting> layerSettings)
        {
            StringBuilder optionsBuilder = new StringBuilder("FLAT PATTERN DXF?");
            StringBuilder invisibleLayersBuilder = new StringBuilder();

            foreach (var layer in layerSettings)
            {
                if (layer.HasVisibilityOption && !layer.IsVisible)
                {
                    // Добавляем дефолтное имя слоя в InvisibleLayers, если он не видим
                    invisibleLayersBuilder.Append($"{layer.LayerName};");
                    continue; // Пропускаем дальнейшую обработку
                }

                // Если есть кастомное имя, оно добавляется в строку
                if (!string.IsNullOrWhiteSpace(layer.CustomName))
                {
                    optionsBuilder.Append($"&{layer.DisplayName}={layer.CustomName}");
                }

                // Если тип линии отличается от дефолтного, добавляем его
                if (layer.SelectedLineType != "Default")
                {
                    optionsBuilder.Append($"&{layer.DisplayName}LineType={GetLineTypeValue(layer.SelectedLineType)}");
                }

                // Если цвет отличается от белого, добавляем его
                if (layer.SelectedColor != "White")
                {
                    optionsBuilder.Append($"&{layer.DisplayName}Color={GetColorValue(layer.SelectedColor)}");
                }
            }

            // Добавляем InvisibleLayers в конец, если есть невидимые слои
            if (invisibleLayersBuilder.Length > 0)
            {
                invisibleLayersBuilder.Length -= 1; // Убираем последний символ ";"
                optionsBuilder.Append($"&InvisibleLayers={invisibleLayersBuilder}");
            }

            return optionsBuilder.ToString();
        }

        // Метод для инициализации настроек слоев
        public static ObservableCollection<LayerSetting> InitializeLayerSettings()
        {
            var layerSettings = new ObservableCollection<LayerSetting>();

            // Добавляем настройки для всех слоев
            layerSettings.Add(new LayerSetting("BendUpLayer", "IV_BEND"));
            layerSettings.Add(new LayerSetting("BendDownLayer", "IV_BEND_DOWN"));
            layerSettings.Add(new LayerSetting("ToolCenterUpLayer", "IV_TOOL_CENTER"));
            layerSettings.Add(new LayerSetting("ToolCenterDownLayer", "IV_TOOL_CENTER_DOWN"));
            layerSettings.Add(new LayerSetting("ArcCentersLayer", "IV_ARC_CENTERS"));
            layerSettings.Add(new LayerSetting("OuterProfileLayer", "IV_OUTER_PROFILE", hasVisibilityOption: false));
            layerSettings.Add(new LayerSetting("InteriorProfilesLayer", "IV_INTERIOR_PROFILES"));
            layerSettings.Add(new LayerSetting("FeatureProfilesUpLayer", "IV_FEATURE_PROFILES"));
            layerSettings.Add(new LayerSetting("FeatureProfilesDownLayer", "IV_FEATURE_PROFILES_DOWN"));
            layerSettings.Add(new LayerSetting("AltRepFrontLayer", "IV_ALTREP_FRONT"));
            layerSettings.Add(new LayerSetting("AltRepBackLayer", "IV_ALTREP_BACK"));
            layerSettings.Add(new LayerSetting("UnconsumedSketchesLayer", "IV_UNCONSUMED_SKETCHES"));
            layerSettings.Add(new LayerSetting("TangentLayer", "IV_TANGENT"));
            layerSettings.Add(new LayerSetting("TangentRollLinesLayer", "IV_ROLL_TANGENT"));
            layerSettings.Add(new LayerSetting("RollLinesLayer", "IV_ROLL"));
            layerSettings.Add(new LayerSetting("UnconsumedSketchConstructionLayer", "IV_UNCONSUMED_SKETCH_CONSTRUCTION"));

            return layerSettings;
        }

        // Доступные цвета
        public static ObservableCollection<string> GetAvailableColors()
        {
            return new ObservableCollection<string>
            {
                "White",
                "Red",
                "Orange",
                "Yellow",
                "Green",
                "Cyan",
                "Blue",
                "Purple",
                "DarkGray",
                "LightGray"
            };
        }

        // Доступные типы линий
        public static ObservableCollection<string> GetLineTypes()
        {
            return new ObservableCollection<string>
            {
                "Chain",
                "Continuous",
                "Dash dotted",
                "Dashed double dotted",
                "Dashed hidden",
                "Dashed",
                "Dashed triple dotted",
                "Default",
                "Dotted",
                "Double dash double dotted",
                "Double dashed",
                "Double dashed dotted",
                "Double dash triple dotted",
                "Long dash dotted",
                "Long dashed double dotted",
                "Long dash triple dotted"
            };
        }
    }
}