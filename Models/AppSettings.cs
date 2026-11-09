// AccessExporter/Models/AppSettings.cs
namespace AccessExporter.Models;

public class AppSettings
{
    public string MdbPath { get; set; } = string.Empty;
    public string OutputFolder { get; set; } = string.Empty;
}