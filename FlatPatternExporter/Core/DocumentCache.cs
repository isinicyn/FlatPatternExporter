using Inventor;

namespace FlatPatternExporter.Core;

public class DocumentCache
{
    private readonly Dictionary<string, PartDocument> _documentCache = [];
    private readonly Dictionary<string, string> _partNumberToFullFileName = [];

    public void AddDocumentToCache(PartDocument partDoc, string partNumber)
    {
        if (!string.IsNullOrEmpty(partNumber) && !_documentCache.ContainsKey(partNumber))
        {
            _documentCache[partNumber] = partDoc;
            _partNumberToFullFileName[partNumber] = partDoc.FullFileName;
        }
    }

    public PartDocument? GetCachedPartDocument(string partNumber)
    {
        return _documentCache.TryGetValue(partNumber, out var partDoc) ? partDoc : null;
    }

    public string? GetCachedPartPath(string partNumber)
    {
        return _partNumberToFullFileName.TryGetValue(partNumber, out var path) ? path : null;
    }

    public void ClearCache()
    {
        _documentCache.Clear();
        _partNumberToFullFileName.Clear();
    }
}