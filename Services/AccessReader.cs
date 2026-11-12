// AccessExporter/Services/AccessReader.cs
using System.Data;
using System.Data.OleDb;
using AccessExporter.Models;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace AccessExporter.Services;

public interface IAccessReader
{
    IReadOnlyList<TableInfo> GetTables(string mdbPath);
    DataTable GetPreview(string mdbPath, ExportItem item, int maxRows);
}

public class AccessReader : IAccessReader
{
    public static string BuildConnectionString(string mdbPath) =>
        $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbPath};Mode=Read;";

    // Shared by preview and export: turns an ExportItem into a ready-to-run command
    public static OleDbCommand CreateCommand(OleDbConnection connection, ExportItem item)
    {
        var command = new OleDbCommand(item.CommandText, connection)
        {
            CommandType = item.CommandType
        };

        // OleDb binds parameters by POSITION, so add them in order; the name is ignored
        foreach (var p in item.Parameters)
        {
            command.Parameters.Add(new OleDbParameter(p.Name, p.Type) { Value = p.Value });
        }

        return command;
    }

    public IReadOnlyList<TableInfo> GetTables(string mdbPath)
    {
        using var connection = new OleDbConnection(BuildConnectionString(mdbPath));
        connection.Open();

        var result = new List<TableInfo>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var noParameters = Array.Empty<QueryParameter>();
        var linkPaths = LoadLinkPaths(mdbPath);

        // 1) Tables, linked tables, select queries without parameters
        var schema = connection.GetSchema("Tables");

        foreach (DataRow row in schema.Rows)
        {
            var name = row["TABLE_NAME"]?.ToString() ?? string.Empty;
            var type = row["TABLE_TYPE"]?.ToString() ?? string.Empty;

            if (IsHidden(name)) continue;

            TableKind? kind = type.ToUpperInvariant() switch
            {
                "TABLE" => TableKind.Table,
                "LINK" => TableKind.LinkTable,
                "VIEW" or "PASS-THROUGH" => TableKind.ViewOrOdbc,
                _ => null
            };

            if (kind is not null && seen.Add(name))
            {
                linkPaths.TryGetValue(name, out var linkPath);
                result.Add(new TableInfo(name, kind.Value, noParameters, linkPath ?? string.Empty));
            }
        }

        // 2) Saved queries that Jet reports as "procedures" (parameter queries, action queries).
        // Optional: if the provider cannot list them, we just show what we have.
        try
        {
            var parameters = LoadProcedureParameters(connection);
            var procedures = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, null);

            if (procedures is not null)
            {
                foreach (DataRow row in procedures.Rows)
                {
                    var name = row["PROCEDURE_NAME"]?.ToString() ?? string.Empty;
                    if (IsHidden(name) || !seen.Add(name)) continue;

                    parameters.TryGetValue(name, out var list);
                    result.Add(new TableInfo(
                        name,
                        TableKind.Procedure,
                        (IReadOnlyList<QueryParameter>?)list ?? noParameters));
                }
            }
        }
        catch
        {
            // ignore
        }

        return result.OrderBy(t => t.Name).ToList();
    }

    public DataTable GetPreview(string mdbPath, ExportItem item, int maxRows)
    {
        using var connection = new OleDbConnection(BuildConnectionString(mdbPath));
        connection.Open();

        using var command = CreateCommand(connection, item);
        using var adapter = new OleDbDataAdapter(command);

        var table = new DataTable();
        adapter.Fill(0, maxRows, table);   // reads at most maxRows rows
        return table;
    }

    private static bool IsHidden(string name) =>
        name.Length == 0
        || name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith("~");

    private static Dictionary<string, List<QueryParameter>> LoadProcedureParameters(OleDbConnection connection)
    {
        var result = new Dictionary<string, List<QueryParameter>>(StringComparer.OrdinalIgnoreCase);

        var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedure_Parameters, null);
        if (schema is null) return result;

        foreach (DataRow row in schema.Rows)
        {
            var procedure = row["PROCEDURE_NAME"]?.ToString() ?? string.Empty;
            var name = row["PARAMETER_NAME"]?.ToString() ?? string.Empty;
            var position = Convert.ToInt32(row["ORDINAL_POSITION"]);

            // OleDbType values mirror the OLE DB DBTYPE numbers, so a plain cast works
            var type = (OleDbType)Convert.ToInt32(row["DATA_TYPE"]);

            if (!result.TryGetValue(procedure, out var list))
            {
                list = new List<QueryParameter>();
                result[procedure] = list;
            }

            list.Add(new QueryParameter(name, position, type));
        }

        foreach (var list in result.Values)
        {
            list.Sort((a, b) => a.Position.CompareTo(b.Position));
        }

        return result;
    }
    // Link table name -> where it points to. Uses ADOX through late binding (no COM reference needed).
    // Any failure just returns what was collected so far: the path is a nice-to-have, never required.
    private static Dictionary<string, string> LoadLinkPaths(string mdbPath)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        object? catalogObject = null;

        try
        {
            var type = Type.GetTypeFromProgID("ADOX.Catalog");
            if (type is null) return result;

            catalogObject = Activator.CreateInstance(type);
            dynamic catalog = catalogObject!;
            catalog.ActiveConnection = BuildConnectionString(mdbPath);

            foreach (dynamic table in catalog.Tables)
            {
                string name = table.Name;
                string tableType = table.Type;
                if (!string.Equals(tableType, "LINK", StringComparison.OrdinalIgnoreCase)) continue;

                var source = ReadProperty(table, "Jet OLEDB:Link Datasource");
                if (source.Length == 0)
                {
                    // ODBC links have no file path: fall back to the connection string
                    source = MaskSecrets(ReadProperty(table, "Jet OLEDB:Link Provider String"));
                }

                var remote = ReadProperty(table, "Jet OLEDB:Remote Table Name");
                if (remote.Length > 0 && !string.Equals(remote, name, StringComparison.OrdinalIgnoreCase))
                {
                    source = $"{source}  ->  [{remote}]";
                }

                result[name] = source;
            }
        }
        catch
        {
            // ADOX missing or not readable: keep going without link paths
        }
        finally
        {
            if (catalogObject is not null) Marshal.ReleaseComObject(catalogObject);
        }

        return result;
    }

    private static string ReadProperty(dynamic table, string propertyName)
    {
        try
        {
            object? value = table.Properties[propertyName].Value;
            return value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    // Connection strings of ODBC links may contain passwords; never show them
    private static string MaskSecrets(string text) =>
        Regex.Replace(text, @"(PWD|Password)\s*=\s*[^;]*", "$1=****", RegexOptions.IgnoreCase);
}