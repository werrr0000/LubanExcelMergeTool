using System.Windows;

namespace LubanExcelMerge.Gui;

public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        var viewModel = new MainWindowViewModel();
        var window = new MainWindow { DataContext = viewModel };
        MainWindow = window;
        window.Show();

        if (e.Args.Length > 0)
            await viewModel.LoadArgumentsAsync(e.Args);
    }
}
