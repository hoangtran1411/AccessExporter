// AccessExporter/Services/ExportService.cs
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Text;
using AccessExporter.Models;
using ClosedXML.Excel;

namespace AccessExporter.Services;

public interface IExportService
{
    Task<IReadOnlyList<string>> ExportAsync(
        ExportOptions options,
        IProgress<ExportProgress>? progress,
        CancellationToken ct);
}

public class ExportService : IExportService
{
    public Task<IReadOnlyList<string>> ExportAsync(
        ExportOptions options,
        IProgress<ExportProgress>? progress,
        CancellationToken ct)
        => Task.Run(() => Export(options, progress, ct), ct);

    private static IReadOnlyList<string> Export(
        ExportOptions options,
        IProgress<ExportProgress>? progress,
        CancellationToken ct)
    {
        Directory.CreateDirectory(options.OutputFolder);

        var files = new List<string>();
        var total = options.Items.Count;

        using var connection = new OleDbConnection(AccessReader.BuildConnectionString(options.MdbPath));
        connection.Open();

        if (options.Format == ExportFormat.Xlsx && options.SingleFile)
        {
            // One workbook, one sheet per item
            using var workbook = new XLWorkbook();
            var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();
                var item = options.Items[i];
                progress?.Report(new ExportProgress(item.Name, i, total));

                var data = LoadTable(connection, item);
                var sheet = workbook.Worksheets.Add(MakeSheetName(item.Name, usedSheetNames));
                FillSheet(sheet, data);
            }

            var path = Path.Combine(
                options.OutputFolder,
                SanitizeFileName(Path.GetFileNameWithoutExtension(options.MdbPath)) + ".xlsx");
            workbook.SaveAs(path);
            files.Add(path);
        }
        else
        {
            // One file per item
            for (var i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();
                var item = options.Items[i];
                progress?.Report(new ExportProgress(item.Name, i, total));

                var extension = options.Format == ExportFormat.Xlsx ? ".xlsx" : ".csv";
                var path = Path.Combine(options.OutputFolder, SanitizeFileName(item.Name) + extension);

                if (options.Format == ExportFormat.Xlsx)
                {
                    using var workbook = new XLWorkbook();
                    var sheet = workbook.Worksheets.Add(MakeSheetName(item.Name, new HashSet<string>()));
                    FillSheet(sheet, LoadTable(connection, item));
                    workbook.SaveAs(path);
                }
                else
                {
                    WriteCsv(connection, item, path, ct);
                }

                files.Add(path);
            }
        }

        progress?.Report(new ExportProgress(string.Empty, total, total));
        return files;
    }

    // ---------- XLSX helpers ----------

    private static DataTable LoadTable(OleDbConnection connection, ExportItem item)
    {
        using var command = AccessReader.CreateCommand(connection, item);
        using var reader = command.ExecuteReader();

        var table = new DataTable();
        table.Load(reader);
        return table;
    }

    private static void FillSheet(IXLWorksheet sheet, DataTable data)
    {
        if (data.Rows.Count == 0)
        {
            // Header only
            for (var c = 0; c < data.Columns.Count; c++)
            {
                sheet.Cell(1, c + 1).Value = data.Columns[c].ColumnName;
            }
        }
        else
        {
            sheet.Cell(1, 1).InsertTable(data, false);
        }

        sheet.SheetView.FreezeRows(1);
        sheet.Columns().AdjustToContents(1, 50); // sample first 50 rows only, to stay fast
    }

    // Excel rules: max 31 chars, none of \ / ? * [ ] :, must be unique (case-insensitive)
    private static string MakeSheetName(string tableName, HashSet<string> used)
    {
        var invalid = new[] { '\\', '/', '?', '*', '[', ']', ':' };
        var cleaned = new string(tableName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        if (cleaned.Length == 0) cleaned = "Sheet";
        if (cleaned.Length > 31) cleaned = cleaned[..31];

        var name = cleaned;
        var n = 2;
        while (!used.Add(name))
        {
            var suffix = $"_{n++}";
            name = cleaned[..Math.Min(cleaned.Length, 31 - suffix.Length)] + suffix;
        }
        return name;
    }

    // ---------- CSV helpers ----------

    // Stream row by row, so large results do not need to fit in memory
    private static void WriteCsv(OleDbConnection connection, ExportItem item, string path, CancellationToken ct)
    {
        try
        {
            using var command = AccessReader.CreateCommand(connection, item);
            using var reader = command.ExecuteReader();
            using var writer = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

            var columnCount = reader.FieldCount;
            var buffer = new object[columnCount];

            for (var i = 0; i < columnCount; i++)
            {
                if (i > 0) writer.Write(',');
                writer.Write(CsvEscape(reader.GetName(i)));
            }
            writer.WriteLine();

            while (reader.Read())
            {
                ct.ThrowIfCancellationRequested();
                reader.GetValues(buffer);

                for (var i = 0; i < columnCount; i++)
                {
                    if (i > 0) writer.Write(',');
                    writer.Write(CsvEscape(FormatValue(buffer[i])));
                }
                writer.WriteLine();
            }
        }
        catch
        {
            // Do not leave a half-written file behind (cancel or error)
            TryDelete(path);
            throw;
        }
    }

    private static string FormatValue(object value) => value switch
    {
        DBNull => string.Empty,
        DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        byte[] => "(binary)",
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };

    private static string CsvEscape(string value)
    {
        if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0) return value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    // ---------- Shared helpers ----------

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
    }

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* ignore */ }
    }
}