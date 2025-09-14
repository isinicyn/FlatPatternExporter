using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FlatPatternExporter.UI.Windows;

namespace FlatPatternExporter.Core;

public class TokenElement
{
    public string Name { get; set; } = string.Empty;
    public int StartIndex { get; set; }
    public int EndIndex { get; set; }
    public Border? VisualElement { get; set; }
    public bool IsCustomText { get; set; } = false;
    
    public bool IsUserDefined => !IsCustomText && PropertyMetadataRegistry.IsUserDefinedProperty(Name);

    public string GetDisplayName()
    {
        if (IsCustomText) return Name;
        
        // Сначала ищем в предопределенных свойствах
        var property = PropertyMetadataRegistry.Properties.Values
            .FirstOrDefault(p => p.IsTokenizable && p.TokenName == Name);
        if (property != null) return property.DisplayName;
        
        // Затем ищем в пользовательских свойствах (TokenName содержит префикс UDP_)
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties
            .FirstOrDefault(p => p.IsTokenizable && p.TokenName == Name);
        if (userProperty != null) return userProperty.DisplayName;
        
        return Name;
    }
}

public class TokenService : INotifyPropertyChanged
{
    private static readonly Regex TokenRegex = new(@"\{(\w+)\}", RegexOptions.Compiled);
    private static readonly Regex CustomTextRegex = new(@"\{CUSTOM:([^}]+)\}", RegexOptions.Compiled);
    private Dictionary<string, Func<PartData, string>> _tokenResolvers;
    private IList<PartData> _partsData;
    private List<TokenElement> _tokenElements = new();
    private WrapPanel? _tokenContainer;

    private string _fileNameTemplate = string.Empty;
    private string _fileNamePreview = string.Empty;
    private bool _isFileNameTemplateValid = true;

    public event PropertyChangedEventHandler? PropertyChanged;

    public TokenService()
    {
        _partsData = new List<PartData>();
        _tokenResolvers = BuildTokenResolvers();
        
        // Подписываемся на изменения в коллекции пользовательских свойств
        PropertyMetadataRegistry.UserDefinedProperties.CollectionChanged += (s, e) =>
        {
            RefreshTokenResolvers();
        };
    }
    
    /// <summary>
    /// Обновляет резолверы токенов при изменении пользовательских свойств
    /// </summary>
    public void RefreshTokenResolvers()
    {
        _tokenResolvers = BuildTokenResolvers();
        UpdatePreview();
    }

    private Dictionary<string, Func<PartData, string>> BuildTokenResolvers()
    {
        var resolvers = new Dictionary<string, Func<PartData, string>>();
        var partDataType = typeof(PartData);
        
        foreach (var prop in PropertyMetadataRegistry.GetTokenizableProperties())
        {
            var tokenName = prop.TokenName;
            if (string.IsNullOrEmpty(tokenName)) continue;
            
            // Для User Defined Properties используем другой подход
            if (prop.Type == PropertyMetadataRegistry.PropertyType.UserDefined)
            {
                // User Defined Properties хранятся в словаре UserDefinedProperties
                resolvers[tokenName] = partData =>
                {
                    if (partData.UserDefinedProperties != null && 
                        partData.UserDefinedProperties.TryGetValue(prop.ColumnHeader, out var value))
                    {
                        return value?.ToString() ?? "";
                    }
                    return "";
                };
            }
            else
            {
                // Находим свойство в PartData
                var property = partDataType.GetProperty(prop.InternalName);
                if (property == null || !property.CanRead) continue;
                
                // Создаем резолвер
                resolvers[tokenName] = partData =>
                {
                    var value = property.GetValue(partData);
                    return PropertyMetadataRegistry.FormatValue(prop.InternalName, value);
                };
            }
        }
        
        return resolvers;
    }

    public string ResolveTemplate(string template, PartData? partData)
    {
        if (string.IsNullOrEmpty(template))
            return string.Empty;

        var tokenCache = new Dictionary<string, string>();
        var result = template;
        
        // Если нет данных детали, возвращаем шаблон с placeholder'ами
        bool usePlaceholders = partData == null || string.IsNullOrEmpty(partData.PartNumber);

        // Обрабатываем обычные токены
        result = TokenRegex.Replace(result, match =>
        {
            var token = match.Groups[1].Value;
            
            if (!tokenCache.TryGetValue(token, out var value))
            {
                if (usePlaceholders)
                {
                    // Если нет данных, возвращаем placeholder
                    var prop = PropertyMetadataRegistry.Properties.Values
                        .FirstOrDefault(p => p.IsTokenizable && p.TokenName == token) ??
                        PropertyMetadataRegistry.UserDefinedProperties
                        .FirstOrDefault(p => p.IsTokenizable && p.TokenName == token);
                    value = prop?.PlaceholderValue ?? $"{{{token}}}";
                }
                else if (_tokenResolvers.TryGetValue(token, out var resolver))
                {
                    value = resolver(partData!);
                    tokenCache[token] = value;
                }
                else
                {
                    value = string.Empty;
                    tokenCache[token] = value;
                }
            }
            
            return value;
        });

        // Обрабатываем кастомные токены
        result = CustomTextRegex.Replace(result, match =>
        {
            return match.Groups[1].Value; // Возвращаем текст как есть
        });

        return SanitizeFileName(result);
    }

    public bool ValidateTemplate(string template)
    {
        if (string.IsNullOrEmpty(template))
            return false;

        // Проверяем обычные токены
        var tokenMatches = TokenRegex.Matches(template);
        var validTokens = tokenMatches.All(match => _tokenResolvers.ContainsKey(match.Groups[1].Value));
        
        // Кастомные токены всегда валидны (не нуждаются в разрешении)
        return validTokens;
    }


    public string PreviewTemplate(string template, PartData? sampleData)
    {
        try
        {
            return ResolveTemplate(template, sampleData);
        }
        catch
        {
            return "Ошибка в шаблоне";
        }
    }

    public (string preview, bool isValid) UpdateFileNamePreview(string fileNameTemplate, IList<PartData> partsData, PartData? selectedData = null)
    {
        if (string.IsNullOrEmpty(fileNameTemplate))
        {
            return ("Не задан шаблон - при экспорте будет использовано значение по умолчанию \"Обозначение детали\"", true);
        }

        var isValid = ValidateTemplate(fileNameTemplate);
        if (!isValid)
        {
            return ("Ошибка в шаблоне - неизвестные токены", false);
        }

        var sampleData = CreateSamplePartData(partsData, selectedData);
        var preview = PreviewTemplate(fileNameTemplate, sampleData);
        
        if (preview == "Ошибка в шаблоне")
        {
            return (preview, false);
        }
        
        if (preview.Length > 255)
        {
            return ($"Имя файла слишком длинное ({preview.Length} символов, максимум 255)", false);
        }
        
        return ($"{preview}.dxf", true);
    }

    public PartData? CreateSamplePartData(IList<PartData> partsData, PartData? selectedData = null)
    {
        // Используем выбранную деталь, если она передана
        if (selectedData != null)
        {
            return selectedData;
        }
        
        // Иначе используем данные первой детали из списка, если есть
        if (partsData.Count > 0)
        {
            return partsData[0];
        }

        // Иначе возвращаем null для отображения placeholder'ов
        return null;
    }

    private string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return fileName;

        var invalidChars = System.IO.Path.GetInvalidFileNameChars();
        var result = fileName;
        
        foreach (var invalidChar in invalidChars)
        {
            result = result.Replace(invalidChar, '_');
        }

        return result;
    }

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public string FileNameTemplate
    {
        get => _fileNameTemplate;
        set
        {
            if (_fileNameTemplate != value)
            {
                _fileNameTemplate = value;
                OnPropertyChanged();
                UpdatePreview();
            }
        }
    }

    public string FileNamePreview
    {
        get => _fileNamePreview;
        private set
        {
            if (_fileNamePreview != value)
            {
                _fileNamePreview = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsFileNameTemplateValid
    {
        get => _isFileNameTemplateValid;
        private set
        {
            if (_isFileNameTemplateValid != value)
            {
                _isFileNameTemplateValid = value;
                OnPropertyChanged();
            }
        }
    }

    public void UpdatePartsData(IList<PartData> partsData)
    {
        _partsData = partsData ?? new List<PartData>();
        UpdatePreview();
    }

    public void UpdatePreviewWithSelectedData(PartData? selectedData)
    {
        var (preview, isValid) = UpdateFileNamePreview(_fileNameTemplate, _partsData, selectedData);
        FileNamePreview = preview;
        IsFileNameTemplateValid = isValid;
    }

    private void UpdatePreview()
    {
        var (preview, isValid) = UpdateFileNamePreview(_fileNameTemplate, _partsData);
        FileNamePreview = preview;
        IsFileNameTemplateValid = isValid;
        UpdateVisualContent();
    }

    public void SetTokenContainer(WrapPanel tokenContainer)
    {
        _tokenContainer = tokenContainer;
        UpdateVisualContent();
    }

    private void UpdateVisualContent()
    {
        if (_tokenContainer == null) return;

        _tokenContainer.Children.Clear();
        _tokenElements.Clear();

        if (string.IsNullOrEmpty(_fileNameTemplate)) return;

        var allMatches = new List<(Match match, bool isCustom)>();
        
        foreach (Match match in TokenRegex.Matches(_fileNameTemplate))
            allMatches.Add((match, false));
        
        foreach (Match match in CustomTextRegex.Matches(_fileNameTemplate))
            allMatches.Add((match, true));
        
        allMatches.Sort((a, b) => a.match.Index.CompareTo(b.match.Index));
        
        for (int i = 0; i < allMatches.Count; i++)
        {
            var (match, isCustom) = allMatches[i];
            var tokenElement = new TokenElement
            {
                Name = match.Groups[1].Value,
                StartIndex = match.Index,
                EndIndex = match.Index + match.Length - 1,
                IsCustomText = isCustom
            };
            
            AddTokenElement(tokenElement, i);
            _tokenElements.Add(tokenElement);
        }
    }

    private void AddTokenElement(TokenElement tokenElement, int index)
    {
        if (_tokenContainer == null) return;

        var border = new Border
        {
            Style = _tokenContainer.FindResource("TokenBlockStyle") as Style,
            Tag = tokenElement.IsCustomText ? "CustomText" : (tokenElement.IsUserDefined ? "UserDefined" : null)
        };

        var textBlock = new TextBlock
        {
            Text = tokenElement.GetDisplayName(),
            Style = _tokenContainer.FindResource("TokenTextStyle") as Style
        };

        border.Child = textBlock;
        tokenElement.VisualElement = border;

        border.MouseDown += (s, e) =>
        {
            if (e.RightButton == MouseButtonState.Pressed) RemoveTokenByIndex(index);
        };

        _tokenContainer.Children.Add(border);
    }

    private void RemoveTokenByIndex(int index)
    {
        if (index < 0 || index >= _tokenElements.Count) return;
        
        var tokenElement = _tokenElements[index];
        var tokenText = tokenElement.IsCustomText 
            ? $"{{CUSTOM:{tokenElement.Name}}}"
            : $"{{{tokenElement.Name}}}";
        
        var newTemplate = _fileNameTemplate.Remove(tokenElement.StartIndex, tokenText.Length);
        FileNameTemplate = newTemplate;
    }

    public void AddToken(string tokenName)
    {
        if (string.IsNullOrEmpty(tokenName)) return;
        
        FileNameTemplate += tokenName;
    }

    public void AddCustomText(string customText)
    {
        if (string.IsNullOrEmpty(customText)) return;
        
        var customTextPattern = @"\{CUSTOM:([^}]+)\}$";
        var match = Regex.Match(_fileNameTemplate, customTextPattern);
        
        if (match.Success)
        {
            var existingCustomText = match.Groups[1].Value;
            var combinedText = existingCustomText + customText;
            var combinedToken = $"{{CUSTOM:{combinedText}}}";
            
            FileNameTemplate = _fileNameTemplate.Substring(0, match.Index) + combinedToken;
        }
        else
        {
            var customToken = $"{{CUSTOM:{customText}}}";
            FileNameTemplate += customToken;
        }
    }
}