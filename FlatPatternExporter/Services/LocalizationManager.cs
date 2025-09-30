using System.ComponentModel;
using System.Globalization;
using System.Resources;

namespace FlatPatternExporter.Services;

public class LocalizationManager : INotifyPropertyChanged
{
    private static LocalizationManager? _instance;
    private static readonly object _lock = new();

    private readonly ResourceManager _resourceManager;
    private CultureInfo _currentCulture;

    public static LocalizationManager Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    _instance ??= new LocalizationManager();
                }
            }
            return _instance;
        }
    }

    private LocalizationManager()
    {
        _resourceManager = new ResourceManager("FlatPatternExporter.Resources.Strings", typeof(LocalizationManager).Assembly);
        _currentCulture = CultureInfo.CurrentUICulture;
    }

    public CultureInfo CurrentCulture
    {
        get => _currentCulture;
        set
        {
            if (!Equals(_currentCulture, value))
            {
                _currentCulture = value;
                CultureInfo.CurrentUICulture = value;
                OnPropertyChanged(nameof(CurrentCulture));
                OnLanguageChanged();
            }
        }
    }

    public string GetString(string key)
    {
        return _resourceManager.GetString(key, _currentCulture) ?? key;
    }

    public string GetString(string key, params object[] args)
    {
        var format = GetString(key);
        return string.Format(_currentCulture, format, args);
    }

    public void SetLanguage(string languageCode)
    {
        var culture = string.IsNullOrEmpty(languageCode) ? CultureInfo.InvariantCulture : new CultureInfo(languageCode);
        CurrentCulture = culture;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? LanguageChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected virtual void OnLanguageChanged()
    {
        LanguageChanged?.Invoke(this, EventArgs.Empty);
    }
}

public static class SupportedLanguages
{
    public static readonly LanguageInfo English = new("en-US", "Language_English");
    public static readonly LanguageInfo Russian = new("ru-RU", "Language_Russian");

    public static readonly LanguageInfo[] All = { English, Russian };
}

public class LanguageInfo : INotifyPropertyChanged
{
    private readonly string _localizationKey;

    public string Code { get; }
    public CultureInfo Culture { get; }

    public string DisplayName => !string.IsNullOrEmpty(_localizationKey)
        ? LocalizationManager.Instance.GetString(_localizationKey)
        : Code;

    public LanguageInfo(string code, string localizationKey)
    {
        Code = code;
        _localizationKey = localizationKey;
        Culture = new CultureInfo(code);

        LocalizationManager.Instance.LanguageChanged += OnLanguageChanged;
    }

    private void OnLanguageChanged(object? sender, EventArgs e)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public override string ToString() => DisplayName;
}