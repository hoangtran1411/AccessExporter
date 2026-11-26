# Contributing to AccessExporter

Thank you for contributing to `AccessExporter`! This guide covers the local
development environment, project architecture, coding standards, and
verification steps expected for pull requests.

## Prerequisites

- **Windows** (WPF desktop application targeting `net10.0-windows`)
- **.NET 10 SDK**
- **32-bit (`x86`) target support**: Keep `<PlatformTarget>x86</PlatformTarget>`
  and `<RuntimeIdentifier>win-x86</RuntimeIdentifier>` intact in
  [AccessExporter.csproj](AccessExporter.csproj). `Microsoft.Jet.OLEDB.4.0` is
  only available to 32-bit processes on 64-bit Windows.

## Development Workflow

1. Fork or clone the repository and create a branch from `main`:

   ```powershell
   git clone https://github.com/hoangtran1411/AccessExporter.git
   cd AccessExporter
   git checkout -b feat/your-change
   ```

2. Build and run locally:

   ```powershell
   dotnet build AccessExporter.csproj
   dotnet run --project AccessExporter.csproj
   ```

3. Keep changes focused and use clear, imperative commit messages describing
   the change.

## Architecture & Conventions

The codebase follows a lightweight MVVM layout:

- **`Models/`**: Immutable data records (`TableInfo`, `ExportItem`,
  `ExportOptions`, `ExportProgress`, `QueryParameter`, `ParameterValue`) and
  observable row/dialog items (`TableItem`, `ParameterInput`, `AppSettings`).
- **`Services/`**:
  - [AccessReader.cs](Services/AccessReader.cs): All `.mdb` schema discovery,
    data preview loading, link-path lookup, and query SQL definition
    extraction.
  - [ExportService.cs](Services/ExportService.cs): File output generation for
    `Xlsx`, `Csv`, and `Sql` formats.
  - [ParameterParser.cs](Services/ParameterParser.cs),
    [DialogService.cs](Services/DialogService.cs), and
    [SettingsService.cs](Services/SettingsService.cs): Typed parameter
    conversion, modal prompt abstraction, and JSON persistence.
- **`ViewModels/` & `Views/`**: [MainViewModel.cs](ViewModels/MainViewModel.cs)
  uses `CommunityToolkit.Mvvm` (`[ObservableProperty]`, `[RelayCommand]`) and
  binds to [MainWindow.xaml](MainWindow.xaml) and
  [ParameterDialog.xaml](Views/ParameterDialog.xaml).

### Coding Guidelines

- **Nullable & warnings**: `<Nullable>enable</Nullable>` and
  `<ImplicitUsings>enable</ImplicitUsings>` are enabled. Changes must build with
  `0 Warning(s)` and `0 Error(s)`.
- **Threading & UI responsiveness**: Keep `OleDbConnection`, `OleDbCommand`,
  and `ADOX` COM calls off the UI thread (`Task.Run`), and gather any modal
  dialog inputs on the UI thread before awaiting background work.
- **Cancellation & partial outputs**: Respect `CancellationToken` during
  export loops and clean up incomplete output files on failure or cancellation
  (`TryDelete`).
- **COM resource cleanup**: When using late-bound `ADOX` objects
  (`Type.GetTypeFromProgID("ADOX.Catalog")`), always release COM wrappers in
  `finally` blocks (`Marshal.ReleaseComObject`) and guard optional metadata
  lookups so missing COM providers do not break core functionality.
- **Security & resilience**: Mask credentials in connection strings
  (`MaskSecrets` in [AccessReader.cs](Services/AccessReader.cs)), bracket
  identifiers safely, and keep convenience features (such as `settings.json`
  persistence or error logging) fail-safe.
- **Formatting & comments**: Use 4-space indentation, preserve existing file
  header comments (`// AccessExporter/...`), and comment non-obvious *why*
  decisions rather than restating what the code does.

## Verification Checklist

Before opening a pull request, verify:

1. **Clean build**:

   ```powershell
   dotnet build AccessExporter.csproj
   ```

2. **Release single-file publish** (required if you modify `.csproj` settings,
   NuGet packages, or native/COM interop):

   ```powershell
   dotnet publish AccessExporter.csproj -c Release -r win-x86
   ```

3. **Manual verification in the app** against a real `.mdb` file covering any
   areas you touched:
   - **Load tables**: Verify tables, linked tables, user-created `MSys*` tables,
     views, and saved queries load cleanly while built-in Access system tables
     stay hidden.
   - **Preview & SQL box**: Check row previews, status bar messages (including
     linked-table `Source:`), and `Run SQL` / `Export result`.
   - **Exports**: Test `Xlsx` (multi-file and `Single file`), `Csv`, and `Sql`
     (per-query and `Single file`), plus `Cancel` and `Open output`.
