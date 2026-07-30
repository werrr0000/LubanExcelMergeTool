using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using LubanExcelMerge.Cli;

namespace LubanExcelMerge.Gui;

public partial class MainWindow : Window
{
    private const uint ScClose = 0xF060;
    private const uint MfByCommand = 0x00000000;
    private const uint MfEnabled = 0x00000000;
    private const uint MfGrayed = 0x00000001;

    private bool _synchronizing;
    private SaveProgressWindow? _saveProgressWindow;

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
            oldViewModel.SavingStateChanged -= ViewModel_SavingStateChanged;
        }
        if (e.NewValue is MainWindowViewModel newViewModel)
        {
            newViewModel.ConflictNavigationRequested += ViewModel_ConflictNavigationRequested;
            newViewModel.ExternalMergeCompleted += ViewModel_ExternalMergeCompleted;
            newViewModel.CloseRequested += ViewModel_CloseRequested;
            newViewModel.SavingStateChanged += ViewModel_SavingStateChanged;
            if (newViewModel.IsSaving)
                ViewModel_SavingStateChanged(true);
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
        if (DataContext is not MainWindowViewModel viewModel)
            return;
        if (viewModel.IsSaving)
        {
            e.Cancel = true;
            return;
        }
        if (viewModel is { IsExternalMergeInvocation: true, IsMergeCompleted: false })
            Environment.ExitCode = viewModel.ExternalMergeExitCode;
    }

    private void ViewModel_SavingStateChanged(bool isSaving)
    {
        if (isSaving)
        {
            SetCloseButtonEnabled(false);
            _saveProgressWindow = new SaveProgressWindow { Owner = this };
            _saveProgressWindow.Show();
            IsEnabled = false;
            _saveProgressWindow.Activate();
            return;
        }

        var progressWindow = _saveProgressWindow;
        _saveProgressWindow = null;
        progressWindow?.Complete();
        IsEnabled = true;
        SetCloseButtonEnabled(true);
        Activate();
    }

    private void SetCloseButtonEnabled(bool enabled)
    {
        var windowHandle = new WindowInteropHelper(this).Handle;
        if (windowHandle == nint.Zero)
            return;
        var systemMenu = GetSystemMenu(windowHandle, false);
        if (systemMenu == nint.Zero)
            return;
        _ = EnableMenuItem(systemMenu, ScClose, MfByCommand | (enabled ? MfEnabled : MfGrayed));
        _ = DrawMenuBar(windowHandle);
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

    [DllImport("user32.dll")]
    private static extern nint GetSystemMenu(nint windowHandle, bool revert);

    [DllImport("user32.dll")]
    private static extern uint EnableMenuItem(nint menuHandle, uint itemId, uint enableFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DrawMenuBar(nint windowHandle);
}
