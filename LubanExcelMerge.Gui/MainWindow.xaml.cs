using System.ComponentModel;
using System.Windows;
using LubanExcelMerge.Cli;

namespace LubanExcelMerge.Gui;

public partial class MainWindow : Window
{
    private bool _synchronizing;

    public MainWindow()
    {
        InitializeComponent();
        DataContextChanged += MainWindow_DataContextChanged;
        Closing += MainWindow_Closing;
        foreach (var grid in Grids)
        {
            grid.SelectionChanged += Spreadsheet_SelectionChanged;
            grid.ViewportChanged += Spreadsheet_ViewportChanged;
        }
    }

    private IReadOnlyList<SpreadsheetGrid> Grids =>
        new[] { BaseSpreadsheet, LocalSpreadsheet, RemoteSpreadsheet, MergedSpreadsheet };

    private void MainWindow_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is MainWindowViewModel oldViewModel)
        {
            oldViewModel.ConflictNavigationRequested -= ViewModel_ConflictNavigationRequested;
            oldViewModel.ExternalMergeCompleted -= ViewModel_ExternalMergeCompleted;
            oldViewModel.CloseRequested -= ViewModel_CloseRequested;
        }
        if (e.NewValue is MainWindowViewModel newViewModel)
        {
            newViewModel.ConflictNavigationRequested += ViewModel_ConflictNavigationRequested;
            newViewModel.ExternalMergeCompleted += ViewModel_ExternalMergeCompleted;
            newViewModel.CloseRequested += ViewModel_CloseRequested;
        }
    }

    private static void ViewModel_ExternalMergeCompleted()
    {
        Environment.ExitCode = 0;
        Application.Current.Shutdown(0);
    }

    private void ViewModel_CloseRequested() => Close();

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (DataContext is MainWindowViewModel { IsExternalMergeInvocation: true, IsMergeCompleted: false } viewModel)
            Environment.ExitCode = viewModel.ExternalMergeExitCode;
    }

    private void ViewModel_ConflictNavigationRequested(int rowIndex, int columnIndex) =>
        NavigateAll(rowIndex, columnIndex);

    private void Spreadsheet_SelectionChanged(object? sender, SpreadsheetSelectionEventArgs e)
    {
        if (_synchronizing)
            return;
        _synchronizing = true;
        try
        {
            foreach (var grid in Grids.Where(grid => !ReferenceEquals(grid, sender)))
                grid.ApplySelection(e.Selection);
        }
        finally
        {
            _synchronizing = false;
        }
        if (DataContext is MainWindowViewModel viewModel)
            viewModel.SelectConflicts(e.ConflictIds);
    }

    private void NavigateAll(int rowIndex, int columnIndex)
    {
        _synchronizing = true;
        try
        {
            foreach (var grid in Grids)
                grid.NavigateTo(rowIndex, columnIndex);
        }
        finally
        {
            _synchronizing = false;
        }
    }

    private void Spreadsheet_ViewportChanged(object? sender, SpreadsheetViewportEventArgs e)
    {
        if (_synchronizing)
            return;
        _synchronizing = true;
        try
        {
            foreach (var grid in Grids.Where(grid => !ReferenceEquals(grid, sender)))
                grid.SetScrollOffsets(e.VerticalOffset, e.HorizontalOffset);
        }
        finally
        {
            _synchronizing = false;
        }
    }
}
