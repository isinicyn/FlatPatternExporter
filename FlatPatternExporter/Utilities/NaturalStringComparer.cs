using System.Runtime.InteropServices;

namespace FlatPatternExporter.Utilities;

/// <summary>
/// Compares strings the way Windows Explorer does: digit sequences are compared as numbers ("2" before "10")
/// </summary>
public sealed class NaturalStringComparer : IComparer<string>
{
    public static readonly NaturalStringComparer Instance = new();

    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    private static extern int StrCmpLogicalW(string x, string y);

    public int Compare(string? x, string? y) => StrCmpLogicalW(x ?? "", y ?? "");
}
