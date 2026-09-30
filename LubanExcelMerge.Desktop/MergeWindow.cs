using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using LubanExcelMerge.Cli;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Desktop;

public sealed class MergeWindow : Window
{
    private readonly string[] _args;
    private readonly TextBlock _status = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _detail = new() { TextWrapping = TextWrapping.Wrap };
    private readonly ComboBox _sheets = new() { MinWidth = 200 };
    private readonly ListBox _conflicts = new() { MinHeight = 80 };
    private readonly CheckBox _reviewed = new() { Name = "Reviewed", Content = "我已复核自动合并结果、结构变化和主键推断提示" };
    private readonly Button _save = new() { Content = "保存并标记已解决", Name = "SaveMerge" };
    private readonly Button _cancel = new() { Content = "取消", Name = "CancelMerge" };
    private readonly StackPanel _choices = new() { Orientation = Orientation.Horizontal, Spacing = 8 };
    private readonly DataGrid[] _tables = new DataGrid[4];
    private MergeCommandOptions? _options;
    private PreparedMergeSession? _session;
    private bool _loaded;
    private MergeDiagnosticLogger? _diagnosticLogger;

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
            item.Conflict.ClearResolution();
            Refresh();
        };
        _choices.Children.Add(clear);
        _choices.Margin = new Thickness(0, 10);
        Add(root, _choices, 3);
        var previews = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*"), RowDefinitions = new RowDefinitions("*,*"), ColumnSpacing = 8, RowSpacing = 8 };
        for (var i = 0; i < 4; i++)
        {
            var panel = new DockPanel();
            var title = new TextBlock { Text = ((MergeGridSide)i).ToString().ToUpperInvariant(), FontWeight = FontWeight.Bold, Margin = new Thickness(0, 4) };
            DockPanel.SetDock(title, Dock.Top);
            panel.Children.Add(title);
            _tables[i] = new DataGrid { IsReadOnly = true, AutoGenerateColumns = false, GridLinesVisibility = DataGridGridLinesVisibility.All };
            panel.Children.Add(_tables[i]);
            Grid.SetRow(panel, i / 2);
            Grid.SetColumn(panel, i % 2);
            previews.Children.Add(panel);
        }
        Add(root, previews, 4);
        var footer = new StackPanel { Spacing = 8, Margin = new Thickness(0, 12, 0, 0) };
        footer.Children.Add(new TextBlock { Text = "Mac 首版不调用 Office 重算：公式缓存可能过期，依赖公式的表请在 Excel/WPS 中重算保存后再导出。", TextWrapping = TextWrapping.Wrap });
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
        _sheets.SelectionChanged += (_, _) => Refresh();
        _reviewed.IsCheckedChanged += (_, _) => UpdateEnabled();
        _save.Click += async (_, _) => await SaveAsync();
        _cancel.Click += (_, _) => Close();
        Closing += (_, e) => e.Cancel = IsBusy;
        Closed += (_, _) => _session?.Dispose();
        Opened += async (_, _) => await LoadAsync();
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
        item.Conflict.Resolve(choice);
        _reviewed.IsChecked = false;
        Refresh();
    }

    private void Refresh()
    {
        if (_session is null || _sheets.SelectedIndex < 0) return;
        var sheet = _session.Sheets[_sheets.SelectedIndex];
        var selectedId = SelectedConflict?.Conflict.Id;
        var items = sheet.Conflicts.Select(c => new ConflictItem(sheet.SheetName, c)).ToArray();
        _conflicts.ItemsSource = items;
        _conflicts.SelectedItem = items.FirstOrDefault(i => i.Conflict.Id == selectedId) ?? items.FirstOrDefault();
        for (var i = 0; i < 4; i++)
        {
            var table = sheet.Comparison.CreateTable((MergeGridSide)i);
            var grid = _tables[i];
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridTextColumn { Header = "行", Binding = new Binding("RowHeader"), Width = new DataGridLength(65) });
            for (var column = 0; column < table.ColumnHeaders.Count; column++)
                grid.Columns.Add(new DataGridTextColumn
                {
                    Header = table.ColumnHeaders[column], Binding = new Binding($"Cells[{column}].DisplayValue"),
                    Width = new DataGridLength(150)
                });
            grid.ItemsSource = table.Rows;
        }
        _status.Text = $"{_session.OutputPath}\n共 {_session.Sheets.Count} 张表，未解决冲突 {_session.RemainingConflicts}/{_session.Conflicts.Count}。"
            + (!_session.LogicalTableRegistered ? "\n注意：此表未登记，当前使用推断主键；保存前必须复核。" : "")
            + (_session.RequiresStructuralChangeConfirmation ? "\n结构变化：" + string.Join("；", _session.StructuralChanges) : "");
        UpdateDetail();
        UpdateEnabled();
    }

    private void UpdateDetail()
    {
        _detail.Text = SelectedConflict is { } item
            ? $"{item.SheetName} · {item.Conflict.Conflict.Message}\n\nBASE：{item.Conflict.BaseValue}\n\nLOCAL（当前分支）：{item.Conflict.LocalValue}\n\nREMOTE（待合并分支）：{item.Conflict.RemoteValue}\n\n当前选择：{item.Conflict.SelectedChoice?.ToString() ?? "未解决"}"
            : "本工作表没有冲突，请复核四表中的自动合并结果。";
        UpdateEnabled();
    }

    public async Task SaveAsync()
    {
        if (ResultExitCode == ExitCodes.Success || IsBusy || _session is not { CanSave: true } || _options is null || _reviewed.IsChecked != true) return;
        SetBusy(true, "正在保存、验证并暂存，请勿退出…");
        try
        {
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
        _choices.IsEnabled = !IsBusy && ResultExitCode != ExitCodes.Success && SelectedConflict is not null;
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
