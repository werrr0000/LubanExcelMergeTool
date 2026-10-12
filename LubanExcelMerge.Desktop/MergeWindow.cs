using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Input;
using Avalonia.VisualTree;
using LubanExcelMerge.Cli;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Desktop;

public sealed class MergeWindow : Window
{
    private readonly string[] _args;
    private readonly TextBlock _status = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _detail = new() { TextWrapping = TextWrapping.Wrap };
    private readonly ComboBox _sheets = new() { MinWidth = 200 };
    private readonly ComboBox _filter = new() { MinWidth = 120 };
    private readonly TextBox _search = new() { Watermark = "搜索冲突…", MinWidth = 180 };
    private readonly ListBox _conflicts = new() { MinHeight = 80 };
    private readonly CheckBox _reviewed = new() { Name = "Reviewed", Content = "我已复核自动合并结果、结构变化和主键推断提示" };
    private readonly Button _save = new() { Content = "保存并标记已解决", Name = "SaveMerge" };
    private readonly Button _cancel = new() { Content = "取消", Name = "CancelMerge" };
    private readonly WrapPanel _choices = new() { Orientation = Orientation.Horizontal };
    private readonly DataGrid[] _tables = new DataGrid[4];
    private MergeCommandOptions? _options;
    private PreparedMergeSession? _session;
    private bool _loaded;
    private MergeDiagnosticLogger? _diagnosticLogger;
    private readonly Stack<IReadOnlyDictionary<string, MergeChoice?>> _undo = new();
    private readonly Stack<IReadOnlyDictionary<string, MergeChoice?>> _redo = new();
    private bool _restoringSnapshot;
    private bool _syncingScroll;
    private int _activeGridColumn = -1;
    private readonly TextBlock _selectionHint = new() { Text = "点击列标题选择整列，再使用当前列按钮。" };
    private readonly List<(Button Header, int Column)> _columnHeaders = new();
    private int _activeGridRow = -1;

    public MergeWindow(string[] args)
    {
        _args = args;
        Title = "Luban Excel Merge · Mac 预览版";
        Width = 1360;
        Height = 900;
        MinWidth = 960;
        MinHeight = 640;
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,180,Auto,*,Auto"), Margin = new Thickness(16) };
        Content = root;
        var top = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        top.Children.Add(new TextBlock { Text = "Excel 三方合并", FontSize = 22, FontWeight = FontWeight.Bold });
        top.Children.Add(_sheets);
        top.Children.Add(new TextBlock { Text = "筛选", VerticalAlignment = VerticalAlignment.Center });
        _filter.ItemsSource = new[] { "全部", "未解决", "已解决" };
        _filter.SelectedIndex = 0;
        top.Children.Add(_filter);
        top.Children.Add(_search);
        var previous = new Button { Content = "上一处", Name = "PreviousConflict" };
        var next = new Button { Content = "下一处", Name = "NextConflict" };
        previous.Click += (_, _) => NavigateConflict(-1);
        next.Click += (_, _) => NavigateConflict(1);
        top.Children.Add(previous);
        top.Children.Add(next);
        Add(root, top, 0);
        _status.Margin = new Thickness(0, 12);
        Add(root, _status, 1);
        var review = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*"), ColumnSpacing = 16 };
        review.Children.Add(_conflicts);
        var details = new ScrollViewer { Content = _detail };
        Grid.SetColumn(details, 1);
        review.Children.Add(details);
        Add(root, review, 2);
        foreach (var choice in Enum.GetValues<MergeChoice>())
        {
            var button = new Button { Content = $"选择 {choice.ToString().ToUpperInvariant()}", Name = $"Choose{choice}" };
            button.Click += (_, _) => Resolve(choice);
            _choices.Children.Add(button);
        }
        var clear = new Button { Content = "清除选择" };
        clear.Click += (_, _) =>
        {
            if (SelectedConflict is not { } item || IsBusy) return;
            SaveSnapshot();
            item.Conflict.ClearResolution();
            Refresh();
        };
        _choices.Children.Add(clear);
        var takeAllRemote = new Button { Content = "全部未解决选 REMOTE", Name = "ChooseAllRemote" };
        takeAllRemote.Click += (_, _) => ResolveAll(MergeChoice.Remote);
        var takeAllLocal = new Button { Content = "全部未解决选 LOCAL", Name = "ChooseAllLocal" };
        takeAllLocal.Click += (_, _) => ResolveAll(MergeChoice.Local);
        foreach (var choice in Enum.GetValues<MergeChoice>())
        {
            var rowButton = new Button { Content = $"当前行选 {choice.ToString().ToUpperInvariant()}" };
            rowButton.Click += (_, _) => ResolveGridScope(choice, columnOnly: false);
            var columnButton = new Button { Content = $"当前列选 {choice.ToString().ToUpperInvariant()}", Name = $"ChooseColumn{choice}" };
            columnButton.Click += (_, _) => ResolveGridScope(choice, columnOnly: true);
            _choices.Children.Add(rowButton);
            _choices.Children.Add(columnButton);
        }
        var undo = new Button { Content = "撤销", Name = "Undo" };
        undo.Click += (_, _) => Undo();
        var redo = new Button { Content = "重做", Name = "Redo" };
        redo.Click += (_, _) => Redo();
        _choices.Children.Add(takeAllRemote);
        _choices.Children.Add(takeAllLocal);
        _choices.Children.Add(undo);
        _choices.Children.Add(redo);
        _choices.Margin = new Thickness(0, 10);
        Add(root, _choices, 3);
        var previews = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*"), RowDefinitions = new RowDefinitions("*,*"), ColumnSpacing = 8, RowSpacing = 8 };
        for (var i = 0; i < 4; i++)
        {
            var panel = new DockPanel();
            var title = new TextBlock { Text = ((MergeGridSide)i).ToString().ToUpperInvariant(), FontWeight = FontWeight.Bold, Margin = new Thickness(0, 4) };
            DockPanel.SetDock(title, Dock.Top);
            panel.Children.Add(title);
            _tables[i] = new SynchronizedDataGrid { IsReadOnly = true, AutoGenerateColumns = false, GridLinesVisibility = DataGridGridLinesVisibility.All, CanUserSortColumns = false, CanUserReorderColumns = false };
            var gridIndex = i;
            _tables[i].SelectionChanged += (_, _) =>
            {

                _activeGridRow = _tables[gridIndex].SelectedIndex;
            };
            ((SynchronizedDataGrid)_tables[i]).ViewportChanged += offset => SynchronizeViewport(gridIndex, offset);
            panel.Children.Add(_tables[i]);
            Grid.SetRow(panel, i / 2);
            Grid.SetColumn(panel, i % 2);
            previews.Children.Add(panel);
        }
        Add(root, previews, 4);
        var footer = new StackPanel { Spacing = 8, Margin = new Thickness(0, 12, 0, 0) };
        footer.Children.Add(new TextBlock { Text = "Mac 首版不调用 Office 重算：公式缓存可能过期，依赖公式的表请在 Excel/WPS 中重算保存后再导出。", TextWrapping = TextWrapping.Wrap });
        footer.Children.Add(_selectionHint);
        footer.Children.Add(_reviewed);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        actions.Children.Add(_save);
        actions.Children.Add(_cancel);
        footer.Children.Add(actions);
        Add(root, footer, 5);
        _conflicts.ItemTemplate = new FuncDataTemplate<ConflictItem>((item, _) => new TextBlock
        {
            Text = item?.ToString(), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(4)
        });
        _conflicts.SelectionChanged += (_, _) => UpdateDetail();
        _sheets.SelectionChanged += (_, _) => { _activeGridColumn = -1; _activeGridRow = -1; Refresh(); };
        _filter.SelectionChanged += (_, _) => RefreshConflictList();
        _search.TextChanged += (_, _) => RefreshConflictList();
        _reviewed.IsCheckedChanged += (_, _) => UpdateEnabled();
        _save.Click += async (_, _) => await SaveAsync();
        _cancel.Click += (_, _) => Close();
        Closing += (_, e) => e.Cancel = IsBusy;
        Closed += (_, _) => _session?.Dispose();
        Opened += async (_, _) => await LoadAsync();
        KeyDown += OnKeyDown;
        UpdateEnabled();
    }

    public bool IsBusy { get; private set; }
    public int ResultExitCode { get; private set; } = ExitCodes.UnresolvedConflicts;
    public PreparedMergeSession? Session => _session;
    private ConflictItem? SelectedConflict => _conflicts.SelectedItem as ConflictItem;

    public async Task LoadAsync()
    {
        if (_loaded) return;
        _loaded = true;
        if (_args.Length == 0)
        {
            _status.Text = "请从 Fork 的自定义合并工具启动，或使用 --base、--local、--remote、--output、--repo-root 参数。";
            return;
        }
        SetBusy(true, "正在读取工作簿并计算差异…");
        try
        {
            _options = CommandLineParser.Parse(_args) with { AllowUnregisteredTable = true };
            _diagnosticLogger = MergeDiagnosticLogger.Create(_options.LogPath);
            _diagnosticLogger?.WriteStarted(_options);
            _session = await Task.Run(() =>
            {
                var session = new LubanMergeCoordinator().Prepare(_options);
                session.RetainSaveSource();
                return session;
            });
            _diagnosticLogger?.WritePrepared(_session);
            _sheets.ItemsSource = _session.Sheets.Select(s => s.SheetName).ToArray();
            _sheets.SelectedIndex = 0;
            Refresh();
        }
        catch (Exception exception) { ReportError(exception); }
        finally { SetBusy(false); }
    }

    private void Resolve(MergeChoice choice)
    {
        if (SelectedConflict is not { } item || IsBusy) return;
        SaveSnapshot();
        item.Conflict.Resolve(choice);
        _reviewed.IsChecked = false;
        Refresh();
    }

    private void Refresh()
    {
        if (_session is null || _sheets.SelectedIndex < 0) return;
        var sheet = _session.Sheets[_sheets.SelectedIndex];
        RefreshConflictList(sheet, SelectedConflict?.Conflict.Id);
        _columnHeaders.Clear();
        for (var i = 0; i < 4; i++)
        {
            var table = sheet.Comparison.CreateTable((MergeGridSide)i);
            var grid = _tables[i];
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridTemplateColumn
            {
                Header = "行",
                Width = new DataGridLength(65),
                CellTemplate = new FuncDataTemplate<MergeGridRow>((row, _) =>
                    new TextBlock { Text = row?.RowHeader, Margin = new Thickness(4), FontWeight = FontWeight.SemiBold })
            });
            for (var column = 0; column < table.ColumnHeaders.Count; column++)
            {
                var columnIndex = column;
                grid.Columns.Add(new DataGridTemplateColumn
                {
                    Header = CreateColumnHeader(i, columnIndex, table.ColumnHeaders[columnIndex]),
                    Width = new DataGridLength(150),
                    CellTemplate = new FuncDataTemplate<MergeGridRow>((row, _) =>
                    {
                        var cell = row?.Cells.ElementAtOrDefault(columnIndex);
                        var text = new TextBlock
                        {
                            Text = cell?.DisplayValue,
                            TextTrimming = TextTrimming.CharacterEllipsis
                        };
                        ToolTip.SetTip(text, cell?.ConflictId is null ? cell?.DisplayValue : $"冲突：{cell.ConflictId}");
                        return new Border
                        {
                            Background = CellBackground(cell?.State ?? MergeGridCellState.Normal),
                            Padding = new Thickness(4, 2),
                            Child = text
                        };
                    })
                });
            }
            grid.ItemsSource = table.Rows;
        }
        _status.Text = $"{_session.OutputPath}\n共 {_session.Sheets.Count} 张表，未解决冲突 {_session.RemainingConflicts}/{_session.Conflicts.Count}。"
            + (!_session.LogicalTableRegistered ? "\n注意：此表未登记，当前使用推断主键；保存前必须复核。" : "")
            + (_session.RequiresStructuralChangeConfirmation ? "\n结构变化：" + string.Join("；", _session.StructuralChanges) : "");
        UpdateDetail();
        UpdateEnabled();
    }

    private void RefreshConflictList(PreparedSheetMerge? sheet = null, string? selectedId = null)
    {
        if (_session is null || _sheets.SelectedIndex < 0) return;
        sheet ??= _session.Sheets[_sheets.SelectedIndex];
        selectedId ??= SelectedConflict?.Conflict.Id;
        var filter = _filter.SelectedItem as string ?? "全部";
        var search = _search.Text?.Trim() ?? string.Empty;
        var items = sheet.Conflicts
            .Where(c => filter switch
            {
                "未解决" => !c.IsResolved,
                "已解决" => c.IsResolved,
                _ => true
            })
            .Where(c => string.IsNullOrEmpty(search) ||
                        c.Conflict.Message.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        c.Id.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        c.BaseValue.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        c.LocalValue.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        c.RemoteValue.Contains(search, StringComparison.OrdinalIgnoreCase))
            .Select(c => new ConflictItem(sheet.SheetName, c))
            .ToArray();
        _conflicts.ItemsSource = items;
        _conflicts.SelectedItem = items.FirstOrDefault(i => i.Conflict.Id == selectedId) ?? items.FirstOrDefault();
    }

    private void ResolveAll(MergeChoice choice)
    {
        if (_session is null || IsBusy) return;
        var unresolved = _session.Conflicts.Where(c => !c.IsResolved).ToArray();
        if (unresolved.Length == 0) return;
        SaveSnapshot();
        foreach (var conflict in unresolved)
            conflict.Resolve(choice);
        _reviewed.IsChecked = false;
        Refresh();
    }

    private Button CreateColumnHeader(int grid, int column, string label)
    {
        var header = new Button
        {
            Name = $"ColumnHeader{grid}_{column}", Content = label,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Padding = new Thickness(10, 4),
            Background = column == _activeGridColumn ? Brushes.LightBlue : Brushes.Transparent
        };
        _columnHeaders.Add((header, column));
        header.Click += (_, e) =>
        {
            if (IsBusy || ResultExitCode == ExitCodes.Success) return;
            _activeGridColumn = column;
            foreach (var item in _columnHeaders)
                item.Header.Background = item.Column == column ? Brushes.LightBlue : Brushes.Transparent;
            _selectionHint.Text = $"已选择整列 {label}（当前工作表）；当前列按钮仅处理该列关联的冲突，整行冲突会处理整条记录。";
            e.Handled = true;
        };
        return header;
    }

    private void ResolveGridScope(MergeChoice choice, bool columnOnly)
    {
        if (_session is null || IsBusy || _sheets.SelectedIndex < 0) return;
        var sheet = _session.Sheets[_sheets.SelectedIndex];
        var table = sheet.Comparison.CreateTable(MergeGridSide.Merged);
        var columnIndex = _activeGridColumn;
        var ids = table.Rows
            .Select((row, index) => (row, index))
            .Where(item => columnOnly ? item.row.Cells.ElementAtOrDefault(columnIndex)?.ConflictId is not null
                                      : item.index == _activeGridRow)
            .SelectMany(item => columnOnly
                ? new[] { item.row.Cells[columnIndex].ConflictId }
                : item.row.Cells.Select(cell => cell.ConflictId))
            .Where(id => id is not null)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (ids.Length == 0)
        {
            _selectionHint.Text = "当前范围没有冲突；请选择含冲突的列或行。";
            return;
        }
        SaveSnapshot();
        foreach (var conflict in _session.Conflicts.Where(c => ids.Contains(c.Id, StringComparer.Ordinal)))
            conflict.Resolve(choice);
        _reviewed.IsChecked = false;
        Refresh();
    }

    private void SynchronizeViewport(int source, Vector offset)
    {
        if (_syncingScroll) return;
        _syncingScroll = true;
        try
        {
            for (var i = 0; i < _tables.Length; i++)
                if (i != source) ((SynchronizedDataGrid)_tables[i]).SetViewport(offset);
        }
        finally { _syncingScroll = false; }
    }

    private void SaveSnapshot()
    {
        if (_session is null || _restoringSnapshot) return;
        _undo.Push(_session.Conflicts.ToDictionary(c => c.Id, c => c.SelectedChoice));
        _redo.Clear();
    }

    private void RestoreSnapshot(IReadOnlyDictionary<string, MergeChoice?> snapshot)
    {
        if (_session is null) return;
        _restoringSnapshot = true;
        try
        {
            foreach (var conflict in _session.Conflicts)
            {
                conflict.ClearResolution();
                if (snapshot.TryGetValue(conflict.Id, out var choice) && choice is { } selected)
                    conflict.Resolve(selected);
            }
        }
        finally { _restoringSnapshot = false; }
        _reviewed.IsChecked = false;
        Refresh();
    }

    private void Undo()
    {
        if (_session is null || IsBusy || _undo.Count == 0) return;
        _redo.Push(_session.Conflicts.ToDictionary(c => c.Id, c => c.SelectedChoice));
        RestoreSnapshot(_undo.Pop());
    }

    private void Redo()
    {
        if (_session is null || IsBusy || _redo.Count == 0) return;
        _undo.Push(_session.Conflicts.ToDictionary(c => c.Id, c => c.SelectedChoice));
        RestoreSnapshot(_redo.Pop());
    }

    private void NavigateConflict(int direction)
    {
        if (_conflicts.ItemCount == 0) return;
        var next = _conflicts.SelectedIndex + direction;
        if (next < 0) next = _conflicts.ItemCount - 1;
        if (next >= _conflicts.ItemCount) next = 0;
        _conflicts.SelectedIndex = next;
        if (_conflicts.SelectedItem is { } selected)
            _conflicts.ScrollIntoView(selected);
    }

    private async void OnKeyDown(object? sender, KeyEventArgs e)
    {
        var modifier = OperatingSystem.IsMacOS() ? KeyModifiers.Meta : KeyModifiers.Control;
        if ((e.KeyModifiers & modifier) != modifier) return;
        if (e.Key == Key.Z) { Undo(); e.Handled = true; }
        else if (e.Key == Key.Y) { Redo(); e.Handled = true; }
        else if (e.Key == Key.S) { await SaveAsync(); e.Handled = true; }
        else if (e.Key == Key.F) { _search.Focus(); e.Handled = true; }
    }

    private void UpdateDetail()
    {
        _detail.Text = SelectedConflict is { } item
            ? $"{item.SheetName} · {item.Conflict.Conflict.Message}\n\nBASE：{item.Conflict.BaseValue}\n\nLOCAL（当前分支）：{item.Conflict.LocalValue}\n\nREMOTE（待合并分支）：{item.Conflict.RemoteValue}\n\n当前选择：{item.Conflict.SelectedChoice?.ToString() ?? "未解决"}"
            : "本工作表没有冲突，请复核四表中的自动合并结果。";
        UpdateEnabled();
    }

    private static IBrush CellBackground(MergeGridCellState state) => state switch
    {
        MergeGridCellState.Conflict => new SolidColorBrush(Color.Parse("#FFD9D6")),
        MergeGridCellState.Added => new SolidColorBrush(Color.Parse("#DDF3E4")),
        MergeGridCellState.Deleted => new SolidColorBrush(Color.Parse("#FFF0BF")),
        MergeGridCellState.Metadata => new SolidColorBrush(Color.Parse("#FFE2B8")),
        MergeGridCellState.Modified => new SolidColorBrush(Color.Parse("#E8F2FA")),
        _ => Brushes.Transparent
    };

    public async Task SaveAsync()
    {
        if (ResultExitCode == ExitCodes.Success || IsBusy || _session is not { CanSave: true } || _options is null || _reviewed.IsChecked != true) return;
        SetBusy(true, "正在保存、验证并暂存，请勿退出…");
        try
        {
            if (_session.RequiresStructuralChangeConfirmation &&
                !await ConfirmStructuralChangesAsync(_session.StructuralChanges))
            {
                SetBusy(false, "已取消保存，MERGED 未被修改。");
                return;
            }
            var completed = await Task.Run(() => MergeCompletion.Save(_session, _options.RepositoryRoot, stage: true));
            _diagnosticLogger?.WriteSaved(_session, completed.SaveResult, gitStaged: true);
            ResultExitCode = ExitCodes.Success;
            // 缓存未验证时保留窗口，让用户看到真实公式状态；明确点击后才退出。
            if (completed.SaveResult.RecalculationStatus is WorkbookRecalculationStatus.SourceCachePreservedUnverified
                or WorkbookRecalculationStatus.DeferredAfterRecalculationFailure)
            {
                _status.Text = "已保存并暂存。公式缓存尚未验证，请在 Excel/WPS 中重算并保存后重新暂存，再进行导出。";
                _save.IsVisible = false;
                _cancel.Content = "完成并返回 Fork";
                _choices.IsVisible = false;
                _reviewed.IsEnabled = false;
            }
            else
            {
                SetBusy(false);
                Close();
            }
        }
        catch (Exception exception) { ReportError(exception); }
        finally { SetBusy(false); }
    }

    private async Task<bool> ConfirmStructuralChangesAsync(IReadOnlyList<string> changes)
    {
        var dialog = new Window
        {
            Title = "确认保存结构变化",
            Width = 620,
            Height = 420,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };
        var content = new StackPanel { Margin = new Thickness(20), Spacing = 12 };
        content.Children.Add(new TextBlock
        {
            Text = "本次合并包含字段列删除、修改或移动，可能影响公式引用和下游 Luban 配置。请确认后继续保存。",
            TextWrapping = TextWrapping.Wrap
        });
        content.Children.Add(new ScrollViewer
        {
            Height = 250,
            Content = new TextBlock
            {
                Text = string.Join(Environment.NewLine, changes.Take(30).Select(change => "• " + change)),
                TextWrapping = TextWrapping.Wrap
            }
        });
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8 };
        var cancel = new Button { Content = "取消", IsCancel = true };
        var confirm = new Button { Content = "确认保存", IsDefault = true };
        cancel.Click += (_, _) => dialog.Close(false);
        confirm.Click += (_, _) => dialog.Close(true);
        buttons.Children.Add(cancel);
        buttons.Children.Add(confirm);
        content.Children.Add(buttons);
        dialog.Content = content;
        return await dialog.ShowDialog<bool>(this);
    }

    private void ReportError(Exception exception)
    {
        ResultExitCode = ExitCodes.ForException(exception);
        _diagnosticLogger?.WriteException(exception, ResultExitCode);
        _status.Text = $"操作失败：{exception.Message}\n{exception.InnerException?.Message}";
    }

    private void SetBusy(bool busy, string? status = null)
    {
        IsBusy = busy;
        if (status is not null) _status.Text = status;
        UpdateEnabled();
    }

    private void UpdateEnabled()
    {
        _save.IsEnabled = !IsBusy && _session?.CanSave == true && _reviewed.IsChecked == true;
        _cancel.IsEnabled = !IsBusy;
        _sheets.IsEnabled = !IsBusy && ResultExitCode != ExitCodes.Success;
        _reviewed.IsEnabled = !IsBusy && ResultExitCode != ExitCodes.Success;
        _conflicts.IsEnabled = !IsBusy;
        _choices.IsEnabled = !IsBusy && ResultExitCode != ExitCodes.Success && _session is not null;
    }

    private static void Add(Grid grid, Control child, int row)
    {
        Grid.SetRow(child, row);
        grid.Children.Add(child);
    }

    private sealed record ConflictItem(string SheetName, ResolvableMergeConflict Conflict)
    {
        public override string ToString() => $"{(Conflict.IsResolved ? "✓ " + Conflict.SelectedChoice : "● 未解决")} · {Conflict.Conflict.Message}";
    }
}
