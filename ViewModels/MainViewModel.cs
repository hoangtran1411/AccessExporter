// AccessExporter/ViewModels/MainViewModel.cs
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Data;
using AccessExporter.Models;
using AccessExporter.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;

namespace AccessExporter.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private const int TablePreviewRows = 100;
    private const int SqlPreviewRows = 1000;

    private readonly IAccessReader _reader = new AccessReader();
    private readonly IExportService _exporter = new ExportService();
    private readonly ISettingsService _settings = new SettingsService();
    private readonly IDialogService _dialogs = new DialogService();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(LoadTablesCommand))]
    [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
    [NotifyCanExecuteChangedFor(nameof(RunSqlCommand))]
    private string _mdbPath = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
    private string _outputFolder = string.Empty;

    [ObservableProperty]
    private string _statusMessage = "Select an .mdb file to begin.";

    [ObservableProperty]
    private TableItem? _selectedTable;

    [ObservableProperty]
    private DataView? _previewData;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanChooseSingleFile))]
    private ExportFormat _selectedFormat = ExportFormat.Xlsx;

    [ObservableProperty]
    private bool _singleFile;

    [ObservableProperty]
    private double _progressValue;

    [ObservableProperty]
    private double _progressMax = 1;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(OpenOutputCommand))]
    private IReadOnlyList<string> _lastExportedFiles = Array.Empty<string>();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunSqlCommand))]
    private string _sqlText = string.Empty;

    public ObservableCollection<TableItem> Tables { get; } = new();

    // View over Tables: filters what the ListBox shows without touching the collection itself
    public ICollectionView TablesView { get; }

    public ExportFormat[] Formats { get; } = Enum.GetValues<ExportFormat>();

    // "Single file" only makes sense for xlsx (CSV has no sheets)
    public bool CanChooseSingleFile => SelectedFormat == ExportFormat.Xlsx;

    public MainViewModel()
    {
        var saved = _settings.Load();
        MdbPath = saved.MdbPath;
        OutputFolder = saved.OutputFolder;

        TablesView = CollectionViewSource.GetDefaultView(Tables);
        TablesView.Filter = item =>
            item is TableItem table
            && (string.IsNullOrWhiteSpace(SearchText)
                || table.Name.Contains(SearchText.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    // Generated hook: re-run the filter whenever the search text changes
    partial void OnSearchTextChanged(string value) => TablesView.Refresh();

    // Generated hooks: save whenever either path changes
    partial void OnMdbPathChanged(string value) => SaveSettings();
    partial void OnOutputFolderChanged(string value) => SaveSettings();

    private void SaveSettings() =>
        _settings.Save(new AppSettings { MdbPath = MdbPath, OutputFolder = OutputFolder });

    // ---------- File / tables ----------

    [RelayCommand]
    private void BrowseFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Access database (*.mdb)|*.mdb|All files (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            MdbPath = dialog.FileName;
        }
    }

    [RelayCommand]
    private void BrowseOutputFolder()
    {
        var dialog = new OpenFolderDialog();

        if (dialog.ShowDialog() == true)
        {
            OutputFolder = dialog.FolderName;
        }
    }

    private bool CanLoadTables() => !string.IsNullOrWhiteSpace(MdbPath);

    [RelayCommand(CanExecute = nameof(CanLoadTables))]
    private async Task LoadTablesAsync()
    {
        try
        {
            StatusMessage = "Loading tables...";
            Tables.Clear();
            PreviewData = null;

            // OleDb has no real async API, so run on a background thread to keep the UI responsive
            var items = await Task.Run(() => _reader.GetTables(MdbPath));

            foreach (var item in items)
            {
                Tables.Add(new TableItem(item.Name, item.Kind, item.Parameters));
            }

            StatusMessage = $"Found {Tables.Count} item(s).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
        }
    }

    // ---------- Preview ----------

    // Generated hook: runs automatically whenever SelectedTable changes
    partial void OnSelectedTableChanged(TableItem? value)
    {
        if (value is null)
        {
            PreviewData = null;
            return;
        }

        _ = LoadPreviewAsync(value);
    }

    private async Task LoadPreviewAsync(TableItem table)
    {
        try
        {
            // A parameter query needs values first. Task.Yield() lets the selection change finish
            // before the modal dialog opens, instead of opening it in the middle of a binding update.
            if (table.Kind == TableKind.Procedure && table.Parameters.Count > 0)
            {
                await Task.Yield();
                if (SelectedTable != table) return;
            }

            var item = BuildItem(table);
            if (item is null)
            {
                PreviewData = null;
                StatusMessage = "Preview canceled.";
                return;
            }

            StatusMessage = $"Loading preview of [{table.Name}]...";

            var path = MdbPath;
            var data = await Task.Run(() => _reader.GetPreview(path, item, TablePreviewRows));

            // Ignore stale result if the user already selected another item
            if (SelectedTable != table) return;

            PreviewData = data.DefaultView;
            StatusMessage = $"[{table.Name}]: showing {data.Rows.Count} row(s), {data.Columns.Count} column(s).";
        }
        catch (Exception ex)
        {
            PreviewData = null;
            StatusMessage = $"Error: {ex.Message}";
        }
    }

    // Builds what to run for a list entry. For a parameter query it asks for the values.
    // Returns null if the user canceled the dialog. Must run on the UI thread.
    private ExportItem? BuildItem(TableItem table)
    {
        if (table.Kind != TableKind.Procedure)
        {
            return ExportItem.ForTable(table.Name);
        }

        var values = new List<ParameterValue>();

        if (table.Parameters.Count > 0)
        {
            var inputs = _dialogs.AskParameters($"Parameters of {table.Name}", table.Parameters, table.LastInputs);
            if (inputs is null) return null;

            table.LastInputs = inputs;

            for (var i = 0; i < table.Parameters.Count; i++)
            {
                values.Add(ParameterParser.Parse(table.Parameters[i], inputs[i]));
            }
        }

        return ExportItem.ForProcedure(table.Name, values);
    }

    // ---------- Select all / none ----------

    [RelayCommand]
    private void SelectAll()
    {
        foreach (var table in TablesView.Cast<TableItem>())
        {
            table.IsSelected = true;
        }
    }

    [RelayCommand]
    private void SelectNone()
    {
        foreach (var table in TablesView.Cast<TableItem>())
        {
            table.IsSelected = false;
        }
    }

    // ---------- SQL box ----------

    private bool CanRunSql() =>
        !string.IsNullOrWhiteSpace(MdbPath) && !string.IsNullOrWhiteSpace(SqlText);

    [RelayCommand(CanExecute = nameof(CanRunSql))]
    private async Task RunSqlAsync()
    {
        try
        {
            // The grid is about to show the SQL result, so drop the list selection
            SelectedTable = null;
            StatusMessage = "Running SQL...";

            var path = MdbPath;
            var item = ExportItem.ForSql("SQL", SqlText);
            var data = await Task.Run(() => _reader.GetPreview(path, item, SqlPreviewRows));

            PreviewData = data.DefaultView;

            StatusMessage = data.Rows.Count >= SqlPreviewRows
                ? $"SQL: showing the first {SqlPreviewRows} rows (Export result writes all rows)."
                : $"SQL: {data.Rows.Count} row(s), {data.Columns.Count} column(s).";
        }
        catch (Exception ex)
        {
            PreviewData = null;
            StatusMessage = $"Error: {ex.Message}";
        }
    }

    // ---------- Open output ----------

    private bool CanOpenOutput() => LastExportedFiles.Count > 0;

    [RelayCommand(CanExecute = nameof(CanOpenOutput))]
    private void OpenOutput()
    {
        try
        {
            if (LastExportedFiles.Count == 1 && File.Exists(LastExportedFiles[0]))
            {
                // Open Explorer with the single exported file highlighted
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{LastExportedFiles[0]}\"",
                    UseShellExecute = true
                });
            }
            else if (Directory.Exists(OutputFolder))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = OutputFolder,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Cannot open folder: {ex.Message}";
        }
    }

    // ---------- Export ----------

    private bool CanExport() =>
        !string.IsNullOrWhiteSpace(MdbPath) && !string.IsNullOrWhiteSpace(OutputFolder);

    // One command for two buttons, so they share the same "already running" guard and the same Cancel:
    //   Export button        -> no CommandParameter -> export the ticked items
    //   Export result button -> CommandParameter="sql" -> export the SQL box
    // IncludeCancelCommand generates ExportCancelCommand, which signals the CancellationToken below
    [RelayCommand(CanExecute = nameof(CanExport), IncludeCancelCommand = true)]
    private async Task ExportAsync(string? source, CancellationToken token)
    {
        var fromSql = source == "sql";

        // Collecting items may show the parameter dialog, so it must run on the UI thread.
        // This is still before the first await, so it does.
        List<ExportItem>? items;
        try
        {
            items = CollectExportItems(fromSql);
        }
        catch (FormatException ex)
        {
            StatusMessage = ex.Message;
            return;
        }

        if (items is null) return;

        var options = new ExportOptions(
            MdbPath,
            items,
            OutputFolder,
            SelectedFormat,
            !fromSql && SingleFile && CanChooseSingleFile);

        LastExportedFiles = Array.Empty<string>();
        ProgressMax = items.Count;
        ProgressValue = 0;

        // Created on the UI thread, so the callback below also runs on the UI thread
        var progress = new Progress<ExportProgress>(p =>
        {
            ProgressMax = p.Total;
            ProgressValue = p.Completed;
            StatusMessage = p.Completed < p.Total
                ? $"Exporting [{p.TableName}] ({p.Completed + 1}/{p.Total})..."
                : "Finishing...";
        });

        try
        {
            var files = await _exporter.ExportAsync(options, progress, token);
            LastExportedFiles = files;
            ProgressValue = ProgressMax;
            StatusMessage = $"Done: {files.Count} file(s) written to {OutputFolder}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Export canceled.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
        }
    }

    // Returns null (and sets StatusMessage) when there is nothing to export or the user canceled
    private List<ExportItem>? CollectExportItems(bool fromSql)
    {
        if (fromSql)
        {
            if (string.IsNullOrWhiteSpace(SqlText))
            {
                StatusMessage = "Type a SQL statement first.";
                return null;
            }

            // The timestamp keeps file names unique between runs
            return new List<ExportItem>
            {
                ExportItem.ForSql($"Query_{DateTime.Now:yyyyMMdd_HHmmss}", SqlText)
            };
        }

        var ticked = Tables.Where(t => t.IsSelected).ToList();
        if (ticked.Count == 0)
        {
            StatusMessage = "Tick at least one item to export.";
            return null;
        }

        var items = new List<ExportItem>();
        foreach (var table in ticked)
        {
            var item = BuildItem(table);
            if (item is null)
            {
                StatusMessage = "Export canceled.";
                return null;
            }

            items.Add(item);
        }

        return items;
    }
}