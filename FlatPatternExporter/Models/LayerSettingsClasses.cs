using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using FlatPatternExporter.Services;

namespace FlatPatternExporter.Models;
    /// <summary>
    /// Localizable list item with automatic display name updates on language change
    /// </summary>
    public class LocalizableItem : INotifyPropertyChanged
    {
        private readonly string _key;

        public string Key => _key;

        public string DisplayName => LocalizationManager.Instance.GetString(_key);

        public LocalizableItem(string key)
        {
            _key = key;
            LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
        }

        private void OnLanguageChanged(object? sender, EventArgs e)
        {
            OnPropertyChanged(nameof(DisplayName));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Constants for default layer settings
    /// </summary>
    public static class LayerDefaults
    {
        public const string DefaultCustomName = "";
        public const string DefaultColor = "White";
        public const string DefaultLineType = "Default";
        public const string OuterProfileLayerName = "OuterProfileLayer";
    }

    /// <summary>
    /// Individual layer settings model with change notification support
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
                OnPropertyChanged(nameof(IsOptionsEnabled));
            }
        }

        public bool IsOptionsEnabled => IsChecked;
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
        /// Resets all layer settings to default values
        /// </summary>
        public void ResetSettings()
        {
            CustomName = LayerDefaults.DefaultCustomName;
            SelectedColor = LayerDefaults.DefaultColor;
            SelectedLineType = LayerDefaults.DefaultLineType;
            OnPropertyChanged(nameof(CanReset));
        }

        private bool IsNotDefault() =>
            CustomName != LayerDefaults.DefaultCustomName ||
            SelectedColor != LayerDefaults.DefaultColor ||
            SelectedLineType != LayerDefaults.DefaultLineType;

        /// <summary>
        /// Checks if current settings differ from default values
        /// </summary>
        public bool HasChanges()
        {
            bool defaultIsChecked = DisplayName == LayerDefaults.OuterProfileLayerName;
            return IsNotDefault() || IsChecked != defaultIsChecked;
        }

        public bool CanReset => IsNotDefault();

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Static helper for working with colors, line types, and layer settings
    /// </summary>
    public static class LayerSettingsHelper
    {
        private static readonly Dictionary<string, string> ColorValues = new()
        {
            ["White"] = "255;255;255",
            ["Red"] = "255;0;0",
            ["Orange"] = "255;81;0",
            ["Yellow"] = "234;255;0",
            ["Green"] = "55;255;0",
            ["Cyan"] = "0;195;255",
            ["Blue"] = "13;0;255",
            ["Purple"] = "255;10;169",
            ["DarkGray"] = "169;169;169",
            ["LightGray"] = "211;211;211"
        };

        private static readonly Dictionary<string, string> LineTypeValues = new()
        {
            ["Chain"] = "37644",
            ["Continuous"] = "37633",
            ["Dash dotted"] = "37638",
            ["Dashed double dotted"] = "37645",
            ["Dashed hidden"] = "37641",
            ["Dashed"] = "37634",
            ["Dashed triple dotted"] = "37647",
            ["Default"] = "37648",
            ["Dotted"] = "37636",
            ["Double dash double dotted"] = "37639",
            ["Double dashed"] = "37637",
            ["Double dashed dotted"] = "37646",
            ["Double dash triple dotted"] = "37640",
            ["Long dash dotted"] = "37642",
            ["Long dashed double dotted"] = "37635",
            ["Long dash triple dotted"] = "37643"
        };

        private static readonly List<LayerDefinition> LayerDefinitions =
        [
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
        ];

        /// <summary>
        /// Converts color name to RGB value
        /// </summary>
        public static string GetColorValue(string colorName) =>
            ColorValues.TryGetValue(colorName, out string? value) ? value : ColorValues["White"];

        /// <summary>
        /// Converts line type name to its corresponding value
        /// </summary>
        public static string GetLineTypeValue(string lineTypeName) =>
            LineTypeValues.TryGetValue(lineTypeName, out string? value) ? value : LineTypeValues["Default"];

        /// <summary>
        /// Initializes layer settings collection with preset values
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
        /// Returns collection of available colors
        /// </summary>
        public static ObservableCollection<LocalizableItem> GetAvailableColors() =>
            new(ColorValues.Keys.Select(key => new LocalizableItem(key)));

        /// <summary>
        /// Returns collection of available line types
        /// </summary>
        public static ObservableCollection<LocalizableItem> GetLineTypes() =>
            new(LineTypeValues.Keys.Select(key => new LocalizableItem(key)));

        /// <summary>
        /// Internal class for layer definitions
        /// </summary>
        private readonly record struct LayerDefinition(string DisplayName, string LayerName, bool CanBeHidden = true);
    }

    /// <summary>
    /// Layer name validator with automatic invalid character sanitization
    /// </summary>
    public static class LayerNameValidator
    {
        private static readonly HashSet<char> InvalidCharacters = 
        [ 
            '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=' 
        ];

        private static readonly string ValidationMessage =
            LocalizationManager.Instance.GetString("Validation_InvalidLayerNameCharacters");

        /// <summary>
        /// Cleans string from invalid characters and shows warning when necessary
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
        /// Displays dialog with validation error information
        /// </summary>
        private static void ShowValidationError()
        {
            System.Windows.MessageBox.Show(
                ValidationMessage,
                LocalizationManager.Instance.GetString("MessageBox_InputError"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }