using System.IO;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Core;

public class ConflictPartNumberGroup
{
    public string PartNumber { get; set; } = "";
    public List<ConflictFileInfo> Files { get; set; } = [];
}

public class ConflictFileInfo
{
    public string FileName { get; set; } = "";
    public string ModelState { get; set; } = "";
    public string FilePath { get; set; } = "";
    public string FileNameOnly => Path.GetFileName(FileName);
}

public static class ConflictDataProcessor
{
    public static List<ConflictPartNumberGroup> PrepareConflictData(Dictionary<string, List<PartConflictInfo>> conflictFileDetails)
    {
        return [.. conflictFileDetails.Select(entry => new ConflictPartNumberGroup
        {
            PartNumber = entry.Key,
            Files = [.. entry.Value.Select(conflictInfo => new ConflictFileInfo
            {
                FileName = conflictInfo.FileName,
                ModelState = conflictInfo.ModelState,
                FilePath = conflictInfo.FileName
            }).OrderBy(f => f.FileName)]
        }).OrderBy(g => g.PartNumber)];
    }
}
