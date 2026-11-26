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
    IReadOnlyDictionary<string, string> GetQueryDefinitions(string mdbPath);
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

    public IReadOnlyDictionary<string, string> GetQueryDefinitions(string mdbPath)
    {
        using var connection = new OleDbConnection(BuildConnectionString(mdbPath));
        connection.Open();
        return LoadQueryDefinitions(connection, mdbPath);
    }

    // Collects saved query SQL from OLE DB schema tables first, then falls back to ADOX for
    // any queries whose definitions were not exposed by GetOleDbSchemaTable.
    internal static Dictionary<string, string> LoadQueryDefinitions(OleDbConnection connection, string mdbPath)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        ReadSchemaDefinitions(connection, OleDbSchemaGuid.Views, "TABLE_NAME", "VIEW_DEFINITION", result);
        ReadSchemaDefinitions(connection, OleDbSchemaGuid.Procedures, "PROCEDURE_NAME", "PROCEDURE_DEFINITION", result);
        MergeAdoxDefinitions(mdbPath, result);

        return result;
    }

    private static void ReadSchemaDefinitions(
        OleDbConnection connection,
        Guid schemaGuid,
        string nameColumn,
        string definitionColumn,
        Dictionary<string, string> target)
    {
        try
        {
            var schema = connection.GetOleDbSchemaTable(schemaGuid, null);
            if (schema is null
                || !schema.Columns.Contains(nameColumn)
                || !schema.Columns.Contains(definitionColumn))
            {
                return;
            }

            foreach (DataRow row in schema.Rows)
            {
                var name = row[nameColumn]?.ToString() ?? string.Empty;
                if (IsHidden(name)) continue;

                var sql = CleanSql(row[definitionColumn]?.ToString());
                if (sql.Length > 0 && !target.ContainsKey(name))
                {
                    target[name] = sql;
                }
            }
        }
        catch
        {
            // Schema rowset not supported or unreadable: keep what we have and let ADOX try
        }
    }

    private static void MergeAdoxDefinitions(string mdbPath, Dictionary<string, string> target)
    {
        object? catalogObject = null;

        try
        {
            var type = Type.GetTypeFromProgID("ADOX.Catalog");
            if (type is null) return;

            catalogObject = Activator.CreateInstance(type);
            dynamic catalog = catalogObject!;
            catalog.ActiveConnection = BuildConnectionString(mdbPath);

            MergeAdoxCollection(catalog.Views, target);
            MergeAdoxCollection(catalog.Procedures, target);
        }
        catch
        {
            // ADOX missing or unreadable: keep definitions collected from OLE DB schema
        }
        finally
        {
            if (catalogObject is not null) Marshal.ReleaseComObject(catalogObject);
        }
    }

    private static void MergeAdoxCollection(dynamic items, Dictionary<string, string> target)
    {
        foreach (var entryObject in items)
        {
            object? commandObject = null;
            try
            {
                dynamic entry = entryObject;
                string name = entry.Name;
                if (IsHidden(name) || target.ContainsKey(name)) continue;

                commandObject = entry.Command;
                if (commandObject is null) continue;

                dynamic command = commandObject;
                var sql = CleanSql((string?)command.CommandText);
                if (sql.Length > 0)
                {
                    target[name] = sql;
                }
            }
            catch
            {
                // Skip individual query definitions ADOX cannot read (e.g., some pass-through queries)
            }
            finally
            {
                if (commandObject is not null) Marshal.ReleaseComObject(commandObject);
                if (entryObject is not null && Marshal.IsComObject(entryObject)) Marshal.ReleaseComObject(entryObject);
            }
        }
    }

    private static string CleanSql(string? sql)
    {
        var trimmed = sql?.Trim('\0', ' ', '\t', '\r', '\n') ?? string.Empty;
        if (trimmed.Length == 0) return string.Empty;
        return trimmed.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", Environment.NewLine);
    }

    private static readonly HashSet<string> BuiltInSystemTables = new(StringComparer.OrdinalIgnoreCase)
    {
        "MSysObjects",
        "MSysACEs",
        "MSysQueries",
        "MSysRelationships",
        "MSysDb",
        "MSysAccounts",
        "MSysGroups",
        "MSysUserList",
        "MSysCmdbars",
        "MSysCompactError",
        "MSysIMEXColumns",
        "MSysIMEXSpecs",
        "MSysModules",
        "MSysModules2",
        "MSysNameMap",
        "MSysResources",
        "MSysComplexColumns",
        "MSysAccessObjects",
        "MSysAccessStorage",
        "MSysAccessXML",
        "MSysNavPaneGroups",
        "MSysNavPaneGroupCategories",
        "MSysNavPaneGroupToObjects",
        "MSysNavPaneObjectIDs"
    };

    // Exclude empty names, temporary ~* objects, and built-in Access/Jet metadata tables,
    // while allowing user-created tables whose names start with "MSys" (e.g., MsysThuMuc).
    internal static bool IsHidden(string name) =>
        name.Length == 0
        || name.StartsWith("~")
        || BuiltInSystemTables.Contains(name)
        || name.StartsWith("MSysComplexType_", StringComparison.OrdinalIgnoreCase);

    private static Dictionary<string, List<QueryParameter>> LoadProcedureParameters(OleDbConnection connection)
    {
        var result = new Dictionary<string, List<QueryParameter>>(StringComparer.OrdinalIgnoreCase);

        try
        {
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
        }
        catch
        {
            // Jet 4.0 does not support Procedure_Parameters; still allow listing Procedures without parameter metadata
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