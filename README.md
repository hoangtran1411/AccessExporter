# AccessExporter

A lightweight Windows WPF desktop utility for browsing Microsoft Access
(`.mdb`) databases, previewing tables and queries, running ad-hoc Jet SQL
statements, and exporting data or saved query definitions to `.xlsx`, `.csv`,
or `.sql`.

## Features

- **Database object discovery**:
  - Lists local tables (`Table`), linked tables (`LinkTable`), views / ODBC
    pass-through queries (`ViewOrOdbc`), and saved action, union, or parameter
    queries (`Procedure`) with color-coded icons.
  - Displays linked-table source paths (or masked ODBC connection strings) in
    the status bar during preview.
  - Automatically includes user-created tables and links whose names start with
    `MSys` / `Msys` (e.g., `MsysThuMuc`), while keeping built-in Access/Jet
    system tables (`MSysObjects`, `MSysACEs`, `MSysQueries`,
    `MSysRelationships`, `MSysAccess*`, `MSysNavPane*`, `MSysIMEX*`,
    `MSysComplexType_*`, etc.) and temporary `~*` objects hidden.
- **Search & selection**:
  - Live case-insensitive name filtering (`Search table name`).
  - `Select all` / `Select none` apply directly to the filtered list.
- **Data preview & ad-hoc SQL**:
  - Clicking an item previews up to 100 rows in the grid (with parameter input
    prompts when parameter metadata is available).
  - Run custom Jet SQL statements (`Run SQL` previews up to 1,000 rows;
    `Export result` exports the full result set).
- **Flexible export formats**:
  - **Xlsx**: Export one `.xlsx` file per item or combine all selected items
    into `<DatabaseName>.xlsx` (`Single file`, one worksheet per item) with
    frozen header rows and auto-adjusted column widths.
  - **Csv**: Stream row-by-row into UTF-8 (with BOM) `.csv` files without
    loading entire tables into memory.
  - **Sql**: Extract saved query definitions (`ViewOrOdbc` and `Procedure`)
    into UTF-8 (no BOM) `.sql` files—either one `.sql` file per query or a
    single combined `<DatabaseName>.sql` script (`Single file`)—without
    executing queries or prompting for parameters. Selected `Table` /
    `LinkTable` items are skipped automatically and reported in the status bar.
- **Convenience & safety**:
  - Background execution with progress bar and cancellation support (`Cancel`).
  - `Open output` opens Explorer directly on the exported file (or output
    directory for multi-file exports).
  - Remembers the last `.mdb` path and output folder in `settings.json` beside
    the executable, and logs unhandled errors to `error.log`.

## Requirements

- **OS**: Windows
- **Database driver**: `Microsoft.Jet.OLEDB.4.0` (built into Windows 32-bit
  subsystem; requires running as 32-bit `x86`)
- **Optional COM component**: `ADOX.Catalog` (standard Windows MDAC component,
  used for reading linked-table sources and as a fallback when extracting query
  SQL)
- **Build SDK**: [.NET 10 SDK](https://dotnet.microsoft.com/)
  (`net10.0-windows`)

## Dependencies

Configured in [AccessExporter.csproj](AccessExporter.csproj):

- `ClosedXML` (`0.105.1`)
- `CommunityToolkit.Mvvm` (`8.4.2`)
- `System.Data.OleDb` (`10.0.14`)

## Getting Started

### Clone

```powershell
git clone https://github.com/hoangtran1411/AccessExporter.git
cd AccessExporter
```

### Build and Run

```powershell
dotnet build AccessExporter.csproj
dotnet run --project AccessExporter.csproj
```

### Publish Single-File Executable

[AccessExporter.csproj](AccessExporter.csproj) is pre-configured for a
self-contained, compressed, single-file `win-x86` executable:

```powershell
dotnet publish AccessExporter.csproj -c Release -r win-x86
```

## Usage

1. Click **Browse...** to select an `.mdb` file and **Output folder...** to
   choose the destination directory.
2. Click **Load tables** to scan tables, linked tables, views, and saved
   queries.
3. Click any item to preview its data, or type a Jet SQL query in the top-right
   box and click **Run SQL**.
4. Tick items to export, choose **Xlsx**, **Csv**, or **Sql** (and optionally
   **Single file** for `Xlsx` or `Sql`), then click **Export**.
5. Click **Open output** after export finishes to locate the generated file(s).

## Project Structure

```text
AccessExporter/
├── AccessExporter.csproj
├── App.xaml / App.xaml.cs
├── AssemblyInfo.cs
├── MainWindow.xaml / MainWindow.xaml.cs
├── Assets/
│   └── grow-plant.ico
├── Models/
│   ├── AppSettings.cs
│   ├── ExportModels.cs
│   ├── ParameterInput.cs
│   └── TableItem.cs
├── Services/
│   ├── AccessReader.cs
│   ├── DialogService.cs
│   ├── ExportService.cs
│   ├── ParameterParser.cs
│   └── SettingsService.cs
├── ViewModels/
│   └── MainViewModel.cs
└── Views/
    └── ParameterDialog.xaml / ParameterDialog.xaml.cs
```

## Notes & Known Limitations

- **`.mdb` / Jet 4.0 focus**: Connections use
  `Provider=Microsoft.Jet.OLEDB.4.0;Data Source=...;Mode=Read;`. Newer `.accdb`
  files that require `Microsoft.ACE.OLEDB.12.0` / `16.0` are not configured by
  default.
- **Query parameter metadata under Jet 4.0**: `Microsoft.Jet.OLEDB.4.0` does
  not support `OleDbSchemaGuid.Procedure_Parameters`. Parameterized saved
  queries still appear in the table list and export cleanly to `.sql`, but
  interactive parameter prompting depends on the OLE DB provider returning
  parameter schema metadata.
- **Pass-through / restricted queries**: Some ODBC pass-through or restricted
  queries may not expose their SQL text through Jet OLE DB or `ADOX.Catalog`;
  in `.sql` exports, those items emit an explanatory SQL comment rather than
  failing the batch.

## Contributing

Contributions are welcome. Please see [CONTRIBUTING.md](CONTRIBUTING.md) for
development setup, coding conventions, and verification guidelines.
