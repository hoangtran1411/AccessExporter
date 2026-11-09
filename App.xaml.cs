using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;

namespace AccessExporter;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private static readonly string LogPath =
       Path.Combine(AppContext.BaseDirectory, "error.log");

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        DispatcherUnhandledException += (s, args) =>
        {
            LogError("Dispatcher Exception", args.Exception);

            MessageBox.Show(
                args.Exception.ToString(),
                "Application Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            args.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            if (args.ExceptionObject is Exception ex)
                LogError("Unhandled Exception", ex);
        };

        TaskScheduler.UnobservedTaskException += (s, args) =>
        {
            LogError("Task Exception", args.Exception);
            args.SetObserved();
        };
    }

    private static void LogError(string source, Exception ex)
    {
        try
        {
            File.AppendAllText(
                LogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}\n" +
                $"{ex}\n\n");
        }
        catch
        {
            // Không để lỗi ghi log gây thêm exception.
        }
    }
}

