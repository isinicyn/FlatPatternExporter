using Inventor;

namespace FlatPatternExporter.Core;

public class DocumentCache
{
    private readonly Dictionary<string, PartDocument> _documentCache = new(StringComparer.OrdinalIgnoreCase);

    public void AddDocumentToCache(PartDocument partDoc)
    {
        _documentCache.TryAdd(partDoc.FullDocumentName, partDoc);
    }

    public PartDocument? GetCachedPartDocument(string fullDocumentName)
    {
        return _documentCache.TryGetValue(fullDocumentName, out var partDoc) ? partDoc : null;
    }

    public void ClearCache()
    {
        _documentCache.Clear();
    }
}
