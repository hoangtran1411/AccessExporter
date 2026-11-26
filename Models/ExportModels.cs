// AccessExporter/Models/ExportModels.cs
using System.Data;
using System.Data.OleDb;

namespace AccessExporter.Models;

public enum ExportFormat
{
    Xlsx,
    Csv,
    Sql
}

// A parameter value already converted to its real type
public record ParameterValue(string Name, OleDbType Type, object Value);

// One thing to preview/export: a table, a saved query, or an ad-hoc SQL statement.
// All three boil down to "command text + command type + parameters".
public record ExportItem(
    string Name,                               // used for the file name / sheet name
    string CommandText,
    CommandType CommandType,
    IReadOnlyList<ParameterValue> Parameters,
    bool IsQueryDefinition = false)
{
    public static ExportItem ForTable(string name) =>
        new(name, $"SELECT * FROM [{name.Replace("]", "]]")}]", CommandType.Text, Array.Empty<ParameterValue>());

    public static ExportItem ForSql(string name, string sql) =>
        new(name, sql, CommandType.Text, Array.Empty<ParameterValue>());

    public static ExportItem ForProcedure(string name, IReadOnlyList<ParameterValue> parameters) =>
        new(name, $"[{name}]", CommandType.StoredProcedure, parameters);

    public static ExportItem ForQueryDefinition(string name) =>
        new(name, string.Empty, CommandType.Text, Array.Empty<ParameterValue>(), IsQueryDefinition: true);
}

public record ExportOptions(
    string MdbPath,
    IReadOnlyList<ExportItem> Items,
    string OutputFolder,
    ExportFormat Format,
    bool SingleFile);

// Completed = number of items already finished (0-based index of the item in progress)
public record ExportProgress(string TableName, int Completed, int Total);