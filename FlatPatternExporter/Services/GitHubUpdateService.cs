using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;
using FlatPatternExporter.Models;

namespace FlatPatternExporter.Services;

public class GitHubUpdateService
{
    private const string GitHubApiBaseUrl = "https://api.github.com";
    private const string RepositoryOwner = "isinicyn";
    private const string RepositoryName = "FlatPatternExporter";

    private readonly HttpClient _httpClient;
    public string? LastError { get; private set; }

    public GitHubUpdateService()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("FlatPatternExporter", "2.0"));
        _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        _httpClient.Timeout = TimeSpan.FromSeconds(30);
    }

    public async Task<ReleaseInfo?> GetLatestReleaseAsync()
    {
        LastError = null;

        try
        {
            var url = $"{GitHubApiBaseUrl}/repos/{RepositoryOwner}/{RepositoryName}/releases/latest";
            var response = await _httpClient.GetAsync(url);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();

                if (response.StatusCode == System.Net.HttpStatusCode.Forbidden &&
                    errorContent.Contains("rate limit", StringComparison.OrdinalIgnoreCase))
                {
                    LastError = "RATE_LIMIT_EXCEEDED";
                }
                else
                {
                    LastError = $"HTTP {(int)response.StatusCode}: {response.ReasonPhrase}\nDetails: {errorContent}";
                }

                return null;
            }

            var json = await response.Content.ReadAsStringAsync();

            var githubRelease = JsonSerializer.Deserialize<GitHubRelease>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (githubRelease == null)
            {
                LastError = "Failed to deserialize GitHub release response";
                return null;
            }

            return new ReleaseInfo
            {
                TagName = githubRelease.TagName ?? string.Empty,
                Name = githubRelease.Name ?? string.Empty,
                Body = githubRelease.Body ?? string.Empty,
                PublishedAt = githubRelease.PublishedAt,
                HtmlUrl = githubRelease.HtmlUrl ?? string.Empty,
                IsPrerelease = githubRelease.Prerelease,
                Assets = githubRelease.Assets?.Select(a => new ReleaseAsset
                {
                    Name = a.Name ?? string.Empty,
                    DownloadUrl = a.BrowserDownloadUrl ?? string.Empty,
                    Size = a.Size
                }).ToList() ?? []
            };
        }
        catch (HttpRequestException ex)
        {
            LastError = $"Network error: {ex.Message}";
            return null;
        }
        catch (TaskCanceledException ex)
        {
            LastError = $"Request timeout: {ex.Message}";
            return null;
        }
        catch (JsonException ex)
        {
            LastError = $"JSON parsing error: {ex.Message}";
            return null;
        }
        catch (Exception ex)
        {
            LastError = $"Unexpected error: {ex.GetType().Name} - {ex.Message}";
            return null;
        }
    }

    public async Task<bool> DownloadReleaseAssetAsync(string downloadUrl, string destinationPath, IProgress<double>? progress = null)
    {
        try
        {
            using var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? 0;
            var downloadedBytes = 0L;

            using var contentStream = await response.Content.ReadAsStreamAsync();
            using var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true);

            var buffer = new byte[8192];
            int bytesRead;

            while ((bytesRead = await contentStream.ReadAsync(buffer)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead));
                downloadedBytes += bytesRead;

                if (totalBytes > 0 && progress != null)
                {
                    var progressPercentage = (double)downloadedBytes / totalBytes * 100;
                    progress.Report(progressPercentage);
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string? TagName { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("body")]
        public string? Body { get; set; }

        [JsonPropertyName("published_at")]
        public DateTime PublishedAt { get; set; }

        [JsonPropertyName("html_url")]
        public string? HtmlUrl { get; set; }

        [JsonPropertyName("prerelease")]
        public bool Prerelease { get; set; }

        [JsonPropertyName("assets")]
        public List<GitHubAsset>? Assets { get; set; }
    }

    private class GitHubAsset
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("browser_download_url")]
        public string? BrowserDownloadUrl { get; set; }

        [JsonPropertyName("size")]
        public long Size { get; set; }
    }
}
