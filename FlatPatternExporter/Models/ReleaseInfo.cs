namespace FlatPatternExporter.Models;

public class ReleaseInfo
{
    public string TagName { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public DateTime PublishedAt { get; set; }
    public string HtmlUrl { get; set; } = string.Empty;
    public bool IsPrerelease { get; set; }
    public List<ReleaseAsset> Assets { get; set; } = [];

    public string Version => TagName.TrimStart('v');
}

public class ReleaseAsset
{
    public string Name { get; set; } = string.Empty;
    public string DownloadUrl { get; set; } = string.Empty;
    public long Size { get; set; }
}
