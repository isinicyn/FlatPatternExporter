using System.Collections.Concurrent;
using FlatPatternExporter.UI.Windows;

namespace FlatPatternExporter.Core;

public class ConflictAnalyzer
{
    private readonly ConcurrentDictionary<string, List<PartConflictInfo>> _partNumberTracker = new();
    private Dictionary<string, List<PartConflictInfo>> _conflictFileDetails = [];

    public Dictionary<string, List<PartConflictInfo>> ConflictFileDetails => _conflictFileDetails;
    public int ConflictCount => _conflictFileDetails.Sum(x => x.Value.Count);
    public bool HasConflicts => _conflictFileDetails.Count > 0;

    public void AddPartToTracker(string partNumber, string fileName, string modelState)
    {
        var conflictInfo = new PartConflictInfo
        {
            PartNumber = partNumber,
            FileName = fileName,
            ModelState = modelState
        };

        _partNumberTracker.AddOrUpdate(partNumber,
            [conflictInfo],
            (key, existingList) =>
            {
                if (!existingList.Any(p => p.UniqueId == conflictInfo.UniqueId))
                {
                    existingList.Add(conflictInfo);
                }
                return existingList;
            });
    }

    public async Task AnalyzeConflictsAsync()
    {
        var conflictingPartNumbers = _partNumberTracker
            .Where(p => p.Value.Count > 1)
            .ToDictionary(p => p.Key, p => p.Value);

        _conflictFileDetails = conflictingPartNumbers;

        await Task.CompletedTask;
    }

    public HashSet<string> GetConflictingPartNumbers()
    {
        return _partNumberTracker
            .Where(p => p.Value.Count > 1)
            .Select(p => p.Key)
            .ToHashSet();
    }

    public void FilterConflictingParts(Dictionary<string, int> sheetMetalParts)
    {
        var conflictingPartNumbers = GetConflictingPartNumbers();

        foreach (var conflictingPartNumber in conflictingPartNumbers)
        {
            sheetMetalParts.Remove(conflictingPartNumber);
        }
    }

    public void Clear()
    {
        _partNumberTracker.Clear();
        _conflictFileDetails.Clear();
    }
}