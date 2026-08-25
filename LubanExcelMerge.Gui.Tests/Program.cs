using System.Windows;
using System.Windows.Threading;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Gui;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        var failures = new List<string>();
        Run("sheet replacement ignores stale grid selection", SheetReplacementIgnoresStaleGridSelection, failures);

        foreach (var failure in failures)
            Console.Error.WriteLine(failure);
        Console.WriteLine($"Executed 1 tests: {1 - failures.Count} passed, {failures.Count} failed.");
        return failures.Count == 0 ? 0 : 1;
    }

    private static void Run(string name, Action test, ICollection<string> failures)
    {
        try
        {
            test();
            Console.WriteLine($"PASS {name}");
        }
        catch (Exception exception)
        {
            failures.Add($"FAIL {name}: {exception}");
        }
    }

    private static void SheetReplacementIgnoresStaleGridSelection()
    {
        var application = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        var grid = new SpreadsheetGrid();
        var window = new Window
        {
            Content = grid,
            Width = 640,
            Height = 360,
            Left = -10_000,
            Top = -10_000,
            ShowInTaskbar = false,
            WindowStyle = WindowStyle.ToolWindow
        };

        try
        {
            window.Show();
            grid.Table = CreateTable("Large", 30, 4);
            window.UpdateLayout();
            grid.NavigateTo(25, 3);
            Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background);

            grid.Table = CreateTable("Small", 5, 4);
            window.UpdateLayout();
            Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.Background);
            grid.NavigateTo(4, 3);
        }
        finally
        {
            window.Close();
            application.Shutdown();
        }
    }

    private static MergeGridTable CreateTable(string title, int rowCount, int columnCount)
    {
        var headers = Enumerable.Range(0, columnCount).Select(index => $"C{index}").ToArray();
        var rows = Enumerable.Range(0, rowCount)
            .Select(rowIndex => new MergeGridRow(
                (rowIndex + 1).ToString(),
                rowIndex.ToString(),
                Enumerable.Range(0, columnCount)
                    .Select(columnIndex => new MergeGridCell($"{rowIndex}:{columnIndex}", MergeGridCellState.Normal))
                    .ToArray()))
            .ToArray();
        return new MergeGridTable(title, headers, rows);
    }
}
