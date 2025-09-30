using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace FlatPatternExporter.Services;

public class VersionInfoService
{
    public static string GetApplicationVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        return informationalVersion ?? assembly.GetName().Version?.ToString() ?? "Version unknown";
    }

    public static string GetLastCommitDate()
    {
        try
        {
            string? executingDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            using var process = new Process();
            process.StartInfo.FileName = "git";
            process.StartInfo.Arguments = "log -1 --format=%cd --date=format:\"%d.%m.%Y %H:%M:%S\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.WorkingDirectory = executingDir ?? AppDomain.CurrentDomain.BaseDirectory;

            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            if (!string.IsNullOrWhiteSpace(output))
            {
                return output.Trim();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error retrieving last commit date: {ex.Message}");
        }

        return "Date not determined";
    }
}