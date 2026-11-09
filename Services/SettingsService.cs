// AccessExporter/Services/SettingsService.cs
using System.IO;
using System.Text.Json;
using AccessExporter.Models;

namespace AccessExporter.Services;

public interface ISettingsService
{
    AppSettings Load();
    void Save(AppSettings settings);
}

public class SettingsService : ISettingsService
{
    // AppContext.BaseDirectory = folder containing the running exe
    private static readonly string FilePath =
        Path.Combine(AppContext.BaseDirectory, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(FilePath)) return new AppSettings();

            var json = File.ReadAllText(FilePath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            // Corrupt or unreadable file: fall back to defaults instead of crashing at startup
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        try
        {
            File.WriteAllText(FilePath, JsonSerializer.Serialize(settings, JsonOptions));
        }
        catch
        {
            // Settings are a convenience only (e.g. read-only folder): never break the app over them
        }
    }
}