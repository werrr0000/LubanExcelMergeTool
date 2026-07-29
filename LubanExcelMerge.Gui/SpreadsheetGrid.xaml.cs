using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using LubanExcelMerge.Cli;

namespace LubanExcelMerge.Gui;

public enum SpreadsheetSelectionKind
{
    Cell,
    Row,
    Column
}

public sealed record SpreadsheetSelection(
    SpreadsheetSelectionKind Kind,
    int RowIndex,
    int ColumnIndex);

public sealed class SpreadsheetSelectionEventArgs : EventArgs
{
    public SpreadsheetSelectionEventArgs(SpreadsheetSelection selection, IReadOnlyList<string> conflictIds)
    {
        Selection = selection;
        ConflictIds = conflictIds;
    }

    public SpreadsheetSelection Selection { get; }
    public IReadOnlyList<string> ConflictIds { get; }
}

public sealed class SpreadsheetViewportEventArgs : EventArgs
{
    public SpreadsheetViewportEventArgs(double verticalOffset, double horizontalOffset)
    {
        VerticalOffset = verticalOffset;
        HorizontalOffset = horizontalOffset;
    }

    public double VerticalOffset { get; }
    public double HorizontalOffset { get; }
}

public partial class SpreadsheetGrid : UserControl
{
    public static readonly DependencyProperty TableProperty = DependencyProperty.Register(
        nameof(Table),
        typeof(MergeGridTable),
        typeof(SpreadsheetGrid),
        new PropertyMetadata(null, OnTableChanged));

    private ScrollViewer? _scrollViewer;
    private bool _suppressSelectionEvent;
    private readonly List<SpreadsheetRowItem> _rows = new();
    private IReadOnlyList<string> _columnHeaders = Array.Empty<string>();
    private SpreadsheetSelection? _selection;
    private SpreadsheetRowItem? _scopeSelectedRow;
    private int? _selectedColumnIndex;

    public SpreadsheetGrid() => InitializeComponent();

    public MergeGridTable? Table
    {
        get => (MergeGridTable?)GetValue(TableProperty);
        set => SetValue(TableProperty, value);
    }

    public event EventHandler<SpreadsheetSelectionEventArgs>? SelectionChanged;
    public event EventHandler<SpreadsheetViewportEventArgs>? ViewportChanged;

    public void NavigateTo(int rowIndex, int columnIndex)
        => ApplySelection(new SpreadsheetSelection(SpreadsheetSelectionKind.Cell, rowIndex, columnIndex));

    public void ApplySelection(SpreadsheetSelection selection)
    {
        if (Table is null || Table.Rows.Count == 0 ||
            selection.ColumnIndex < 0 || selection.ColumnIndex >= Grid.Columns.Count)
            return;

        var rowIndex = selection.RowIndex >= 0
            ? selection.RowIndex
            : Grid.CurrentCell.Item is SpreadsheetRowItem current
                ? current.RowIndex
                : 0;
        if (rowIndex < 0 || rowIndex >= _rows.Count)
            rowIndex = 0;
        var columnIndex = selection.ColumnIndex;
        var row = _rows[rowIndex];
        var column = Grid.Columns[columnIndex];
        _selection = selection with { RowIndex = selection.Kind == SpreadsheetSelectionKind.Column ? -1 : rowIndex };
        UpdateRowSelection(selection.Kind == SpreadsheetSelectionKind.Row ? row : null);
        UpdateColumnSelection(selection.Kind == SpreadsheetSelectionKind.Column ? columnIndex : null);

        _suppressSelectionEvent = true;
        Grid.CurrentCell = new DataGridCellInfo(row, column);
        Grid.SelectedCells.Clear();
        Grid.SelectedCells.Add(Grid.CurrentCell);
        Grid.ScrollIntoView(row, column);
        Dispatcher.BeginInvoke(
            () => _suppressSelectionEvent = false,
            DispatcherPriority.Background);
    }

    public void SetScrollOffsets(double verticalOffset, double horizontalOffset)
    {
        if (_scrollViewer is null)
            return;
        _scrollViewer.ScrollToVerticalOffset(verticalOffset);
        _scrollViewer.ScrollToHorizontalOffset(horizontalOffset);
    }

    private static void OnTableChanged(DependencyObject sender, DependencyPropertyChangedEventArgs eventArgs)
    {
        ((SpreadsheetGrid)sender).ApplyTable((MergeGridTable?)eventArgs.NewValue);
    }

    private void ApplyTable(MergeGridTable? table)
    {
        UpdateRowSelection(null);
        UpdateColumnSelection(null);
        if (table is null)
        {
            Grid.ItemsSource = null;
            Grid.Columns.Clear();
            _rows.Clear();
            _columnHeaders = Array.Empty<string>();
            return;
        }

        var headersChanged = !_columnHeaders.SequenceEqual(table.ColumnHeaders, StringComparer.Ordinal);
        _columnHeaders = table.ColumnHeaders;
        if (headersChanged)
        {
            Grid.Columns.Clear();
            for (var columnIndex = 0; columnIndex < table.ColumnHeaders.Count; columnIndex++)
                Grid.Columns.Add(CreateColumn(table.ColumnHeaders[columnIndex], columnIndex));
        }
        _rows.Clear();
        _rows.AddRange(Enumerable.Range(0, table.Rows.Count)
            .Select(rowIndex => new SpreadsheetRowItem(table, rowIndex)));
        Grid.ItemsSource = _rows;
        if (_selection is not null)
            ApplySelection(_selection);
    }

    private DataGridTextColumn CreateColumn(string header, int columnIndex)
    {
        var elementStyle = new Style(typeof(TextBlock));
        elementStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
        elementStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,
            new Binding($"Source.Cells[{columnIndex}].DisplayValue")));
        return new DataGridTextColumn
        {
            Header = CreateColumnHeader(header, false),
            Binding = new Binding($"Source.Cells[{columnIndex}].DisplayValue") { Mode = BindingMode.OneWay },
            Width = new DataGridLength(columnIndex == 0 ? 92 : 112),
            MinWidth = 54,
            CellStyle = CreateCellStyle(columnIndex, false),
            ElementStyle = elementStyle
        };
    }

    private static Style CreateCellStyle(int columnIndex, bool columnSelected)
    {
        var style = new Style(typeof(DataGridCell));
        style.Setters.Add(new Setter(BackgroundProperty, Brushes.White));
        style.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Color.FromRgb(31, 39, 46))));
        style.Setters.Add(new Setter(PaddingProperty, new Thickness(6, 2, 6, 2)));
        style.Setters.Add(new Setter(VerticalContentAlignmentProperty, VerticalAlignment.Center));
        style.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
        AddStateTrigger(style, columnIndex, MergeGridCellState.Modified, "#E8F2FA");
        AddStateTrigger(style, columnIndex, MergeGridCellState.Conflict, "#FFD9D6");
        AddStateTrigger(style, columnIndex, MergeGridCellState.Added, "#DDF3E4");
        AddStateTrigger(style, columnIndex, MergeGridCellState.Deleted, "#FFF0BF");
        AddStateTrigger(style, columnIndex, MergeGridCellState.Metadata, "#FFE2B8");
        if (columnSelected)
        {
            style.Setters.Add(new Setter(BorderBrushProperty, new SolidColorBrush(Color.FromRgb(23, 107, 135))));
            style.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(2, 0, 2, 0)));
        }

        var selected = new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true };
        selected.Setters.Add(new Setter(BorderBrushProperty, new SolidColorBrush(Color.FromRgb(23, 107, 135))));
        selected.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(2)));
        selected.Setters.Add(new Setter(ForegroundProperty, Brushes.Black));
        style.Triggers.Add(selected);
        return style;
    }

    private static void AddStateTrigger(Style style, int columnIndex, MergeGridCellState state, string color)
    {
        var trigger = new DataTrigger
        {
            Binding = new Binding($"Source.Cells[{columnIndex}].State"),
            Value = state
        };
        trigger.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString(color))));
        style.Triggers.Add(trigger);
    }

    private void Grid_Loaded(object sender, RoutedEventArgs e)
    {
        _scrollViewer = FindVisualChild<ScrollViewer>(Grid);
        if (_scrollViewer is not null)
            _scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
    }

    private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e) =>
        ViewportChanged?.Invoke(this, new SpreadsheetViewportEventArgs(e.VerticalOffset, e.HorizontalOffset));

    private void Grid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        if (_suppressSelectionEvent || Table is null || Grid.CurrentCell.Item is not SpreadsheetRowItem row)
            return;
        var rowIndex = row.RowIndex;
        var columnIndex = Grid.CurrentCell.Column?.DisplayIndex ?? -1;
        if (rowIndex < 0 || columnIndex < 0 || columnIndex >= row.Source.Cells.Count)
            return;
        var selection = new SpreadsheetSelection(SpreadsheetSelectionKind.Cell, rowIndex, columnIndex);
        ApplySelection(selection);
        RaiseSelectionChanged(selection);
    }

    private void Grid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (Table is null)
            return;
        if (FindVisualParent<Thumb>(e.OriginalSource as DependencyObject) is not null)
            return;

        var columnHeader = FindVisualParent<DataGridColumnHeader>(e.OriginalSource as DependencyObject);
        if (columnHeader?.Column is not null)
        {
            var selection = new SpreadsheetSelection(
                SpreadsheetSelectionKind.Column,
                -1,
                columnHeader.Column.DisplayIndex);
            ApplySelection(selection);
            RaiseSelectionChanged(selection);
            e.Handled = true;
            return;
        }

        var rowHeader = FindVisualParent<DataGridRowHeader>(e.OriginalSource as DependencyObject);
        if (rowHeader?.DataContext is SpreadsheetRowItem row)
        {
            var rowIndex = row.RowIndex;
            var columnIndex = Grid.CurrentCell.Column?.DisplayIndex ?? 0;
            var selection = new SpreadsheetSelection(SpreadsheetSelectionKind.Row, rowIndex, columnIndex);
            ApplySelection(selection);
            RaiseSelectionChanged(selection);
            e.Handled = true;
        }
    }

    private void RaiseSelectionChanged(SpreadsheetSelection selection)
    {
        if (Table is null)
            return;
        IEnumerable<string?> ids = selection.Kind switch
        {
            SpreadsheetSelectionKind.Row => Table.Rows[selection.RowIndex].Cells.Select(cell => cell.ConflictId),
            SpreadsheetSelectionKind.Column => Table.Rows.Select(row => row.Cells[selection.ColumnIndex].ConflictId),
            _ => new[] { Table.Rows[selection.RowIndex].Cells[selection.ColumnIndex].ConflictId }
        };
        SelectionChanged?.Invoke(this, new SpreadsheetSelectionEventArgs(
            selection,
            ids.Where(id => id is not null).Cast<string>().Distinct(StringComparer.Ordinal).ToArray()));
    }

    private void UpdateRowSelection(SpreadsheetRowItem? row)
    {
        if (ReferenceEquals(_scopeSelectedRow, row))
            return;
        if (_scopeSelectedRow is not null)
            _scopeSelectedRow.IsScopeSelected = false;
        _scopeSelectedRow = row;
        if (_scopeSelectedRow is not null)
            _scopeSelectedRow.IsScopeSelected = true;
    }

    private void UpdateColumnSelection(int? columnIndex)
    {
        if (_selectedColumnIndex == columnIndex)
            return;
        if (_selectedColumnIndex is int previous)
            ApplyColumnSelection(previous, false);
        _selectedColumnIndex = columnIndex;
        if (_selectedColumnIndex is int current)
            ApplyColumnSelection(current, true);
    }

    private void ApplyColumnSelection(int columnIndex, bool selected)
    {
        if (columnIndex < 0 || columnIndex >= Grid.Columns.Count || columnIndex >= _columnHeaders.Count)
            return;
        Grid.Columns[columnIndex].CellStyle = CreateCellStyle(columnIndex, selected);
        Grid.Columns[columnIndex].Header = CreateColumnHeader(_columnHeaders[columnIndex], selected);
    }

    private static Border CreateColumnHeader(string text, bool selected)
    {
        var header = new Border
        {
            Background = selected ? new SolidColorBrush(Color.FromRgb(215, 235, 241)) : Brushes.Transparent,
            Padding = new Thickness(8, 4, 8, 4),
            Child = new TextBlock
            {
                Text = text,
                Foreground = selected ? new SolidColorBrush(Color.FromRgb(23, 107, 135)) :
                    new SolidColorBrush(Color.FromRgb(57, 67, 77)),
                FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center
            }
        };
        return header;
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
        {
            var child = VisualTreeHelper.GetChild(parent, index);
            if (child is T result)
                return result;
            var descendant = FindVisualChild<T>(child);
            if (descendant is not null)
                return descendant;
        }
        return null;
    }

    private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child is not null)
        {
            if (child is T result)
                return result;
            child = VisualTreeHelper.GetParent(child);
        }
        return null;
    }

    private sealed class SpreadsheetRowItem : INotifyPropertyChanged
    {
        private bool _isScopeSelected;

        private readonly MergeGridTable _table;

        public SpreadsheetRowItem(MergeGridTable table, int rowIndex)
        {
            _table = table;
            RowIndex = rowIndex;
        }

        public MergeGridRow Source => _table.Rows[RowIndex];
        public int RowIndex { get; }
        public bool IsScopeSelected
        {
            get => _isScopeSelected;
            set
            {
                if (_isScopeSelected == value)
                    return;
                _isScopeSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsScopeSelected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
