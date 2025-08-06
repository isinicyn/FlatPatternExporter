using System.Text.RegularExpressions;

namespace FlatPatternExporter;

public class TokenService
{
    private static readonly Regex TokenRegex = new(@"\{(\w+)\}", RegexOptions.Compiled);
    private readonly Dictionary<string, Func<PartData, string>> _tokenResolvers;

    public TokenService()
    {
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

        var result = TokenRegex.Replace(template, match =>
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

        return SanitizeFileName(result);
    }

    public bool ValidateTemplate(string template)
    {
        if (string.IsNullOrEmpty(template))
            return false;

        var matches = TokenRegex.Matches(template);
        return matches.All(match => _tokenResolvers.ContainsKey(match.Groups[1].Value));
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

        return result.Length > 255 ? result[..255] : result;
    }
}