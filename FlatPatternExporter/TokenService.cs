using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace FlatPatternExporter;

public class TokenService : INotifyPropertyChanged
{
    private static readonly Regex TokenRegex = new(@"\{(\w+)\}", RegexOptions.Compiled);
    private static readonly Regex CustomTextRegex = new(@"\{CUSTOM:([^}]+)\}", RegexOptions.Compiled);
    private readonly Dictionary<string, Func<PartData, string>> _tokenResolvers;
    private IList<PartData> _partsData;

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
    }
}