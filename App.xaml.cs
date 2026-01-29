using System.Configuration;
using System.Data;
using System.Windows;

namespace ComPortWatcher;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    public App()
    {
        try
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                LogError($"DispatcherUnhandledException: {e.Exception}");
                e.Handled = false;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                LogError($"UnhandledException: {e.ExceptionObject}");
            };
        }
        catch (Exception ex)
        {
            LogError($"App Constructor Error: {ex}");
        }
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            base.OnStartup(e);
            // StartupUri で MainWindow が自動的に作成されるため、ここでは何もしない
        }
        catch (Exception ex)
        {
            LogError($"OnStartup Error: {ex}");
            throw;
        }
    }

    private static void LogError(string message)
    {
        try
        {
            var logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ComPortWatcher_error.log");
            System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\n");
        }
        catch { }
    }
}


