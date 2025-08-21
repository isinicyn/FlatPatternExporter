using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FlatPatternExporter;

public class TokenElement
{
    public string Name { get; set; } = string.Empty;
    public int StartIndex { get; set; }
    public int EndIndex { get; set; }
    public Border? VisualElement { get; set; }
    public bool IsCustomText { get; set; } = false;
}

public class TokenService : INotifyPropertyChanged
{
    private static readonly Regex TokenRegex = new(@"\{(\w+)\}", RegexOptions.Compiled);
    private static readonly Regex CustomTextRegex = new(@"\{CUSTOM:([^}]+)\}", RegexOptions.Compiled);
    private readonly Dictionary<string, Func<PartData, string>> _tokenResolvers;
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
        _tokenResolvers = new Dictionary<string, Func<PartData, string>>
        {
            ["PartNumber"] = partData => partData.PartNumber ?? string.Empty,
            ["Quantity"] = partData => partData.Quantity.ToString(),
            ["Material"] = partData => partData.Material ?? string.Empty,
            ["Thickness"] = partData => partData.Thickness.ToString("F1"),
            ["Description"] = partData => partData.Description ?? string.Empty,
            ["Author"] = partData => partData.Author ?? string.Empty,
            ["Revision"] = partData => partData.Revision ?? string.Empty,
            ["Project"] = partData => partData.Project ?? string.Empty,
            ["Mass"] = partData => partData.Mass ?? string.Empty,
            ["FlatPatternWidth"] = partData => partData.FlatPatternWidth ?? string.Empty,
            ["FlatPatternLength"] = partData => partData.FlatPatternLength ?? string.Empty,
            ["FlatPatternArea"] = partData => partData.FlatPatternArea ?? string.Empty,
        };
    }

    public string ResolveTemplate(string template, PartData partData)
    {
        if (string.IsNullOrEmpty(template) || partData == null)
            return string.Empty;

        var tokenCache = new Dictionary<string, string>();
        var result = template;

        // Обрабатываем обычные токены
        result = TokenRegex.Replace(result, match =>
        {
            var token = match.Groups[1].Value;
            
            if (!tokenCache.TryGetValue(token, out var value))
            {
                if (_tokenResolvers.TryGetValue(token, out var resolver))
                {
                    value = resolver(partData);
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


    public string PreviewTemplate(string template, PartData sampleData)
    {
        if (sampleData == null)
            return "Нет данных для предпросмотра";

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

    public PartData CreateSamplePartData(IList<PartData> partsData, PartData? selectedData = null)
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

        // Иначе возвращаем placeholder данные
        return new PartData
        {
            PartNumber = "{PartNumber}",
            Quantity = 0,
            Material = "{Material}",
            Thickness = 0.0,
            Description = "{Description}",
            Author = "{Author}",
            Revision = "{Revision}",
            Project = "{Project}",
            Mass = "{Mass}",
            FlatPatternWidth = "{FlatPatternWidth}",
            FlatPatternLength = "{FlatPatternLength}",
            FlatPatternArea = "{FlatPatternArea}"
        };
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
            Tag = tokenElement.IsCustomText ? "CustomText" : null
        };

        var textBlock = new TextBlock
        {
            Text = tokenElement.IsCustomText ? $"{tokenElement.Name}" : tokenElement.Name,
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