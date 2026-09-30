using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;
using LubanExcelMerge.Cli;

namespace LubanExcelMerge.Desktop;

public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        if (args is ["--help"] or ["-h"])
        {
            Console.WriteLine(CommandLineParser.Usage);
            return 0;
        }
        if (args.FirstOrDefault() is "git-config" or "diagnostic-package" || args.Contains("--headless"))
            return CliApplication.Run(args, Console.Out, Console.Error);
        return BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<MergeApp>().UsePlatformDetect();
}

public sealed class MergeApp : Application
{
    public override void Initialize()
    {
        RequestedThemeVariant = ThemeVariant.Light;
        Styles.Add(new FluentTheme());
        Styles.Add(new Avalonia.Markup.Xaml.Styling.StyleInclude(new Uri("avares://LubanExcelMerge.Desktop/"))
        {
            Source = new Uri("avares://Avalonia.Controls.DataGrid/Themes/Fluent.xaml")
        });
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var window = new MergeWindow(desktop.Args ?? []);
            desktop.MainWindow = window;
            desktop.ShutdownMode = ShutdownMode.OnMainWindowClose;
            // Cmd+Q 等系统退出入口也必须遵守保存门禁。
            desktop.ShutdownRequested += (_, e) => e.Cancel = window.IsBusy;
            desktop.Exit += (_, e) => e.ApplicationExitCode = window.ResultExitCode;
        }
        base.OnFrameworkInitializationCompleted();
    }
}
