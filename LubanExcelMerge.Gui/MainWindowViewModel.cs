using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Git;
using Microsoft.Win32;

namespace LubanExcelMerge.Gui;

public sealed record RecalculationModeOption(string Value, string DisplayName);

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly LubanMergeCoordinator _coordinator = new();
    private readonly GitMergedFileStager _gitStager = new();
    private readonly Stack<ResolutionBatch> _undo = new();
    private readonly Stack<ResolutionBatch> _redo = new();
    private readonly HashSet<string> _selectedConflictIds = new(StringComparer.Ordinal);
    private readonly Func<IReadOnlyList<string>, bool> _confirmStructuralChanges;
    private MergeDiagnosticLogger? _diagnosticLogger;
    private PreparedMergeSession? _session;
    private SheetTabViewModel? _selectedSheet;
    private ConflictItemViewModel? _selectedConflict;
    private MergeGridTable? _baseGrid;
    private MergeGridTable? _localGrid;
    private MergeGridTable? _remoteGrid;
    private MergeGridTable? _mergedGrid;
    private string _basePath = string.Empty;
    private string _localPath = string.Empty;
    private string _remotePath = string.Empty;
    private string _outputPath = string.Empty;
    private string _repositoryRoot = string.Empty;
    private string _dataRoot = string.Empty;
    private string _tablesPath = string.Empty;
    private string _searchText = string.Empty;
    private string _selectedFilter = "全部";
    private RecalculationModeOption _selectedRecalculationMode;
    private string _statusText = "请选择四个版本文件并加载";
    private bool _isBusy;
    private bool _isSaving;
    private bool _settingSelectionFromGrid;
    private int _processedMergeNavigationIndex = -1;

    public MainWindowViewModel(Func<IReadOnlyList<string>, bool>? confirmStructuralChanges = null)
    {
        _confirmStructuralChanges = confirmStructuralChanges ?? ConfirmStructuralChanges;
        RecalculationModes =
        [
            new RecalculationModeOption("auto", "自动（推荐）"),
            new RecalculationModeOption("always", "始终重算"),
            new RecalculationModeOption("never", "从不（WPS 手动重算）")
        ];
        _selectedRecalculationMode = RecalculationModes[0];
        ConflictsView = CollectionViewSource.GetDefaultView(Conflicts);
        ConflictsView.Filter = FilterConflict;
        LoadCommand = new AsyncRelayCommand(LoadFromFieldsAsync, () => !IsBusy && HasRequiredInputs);
        SaveCommand = new AsyncRelayCommand(SaveAsync, () => !IsBusy && _session?.CanSave == true);
        ResolveCommand = new RelayCommand(ResolveSelected, _ => SelectedConflictCount > 0 && !IsBusy);
        PreviousCommand = new RelayCommand(_ => MoveSelection(-1), _ => _session?.Conflicts.Count > 0);
        NextCommand = new RelayCommand(_ => MoveSelection(1), _ => _session?.Conflicts.Count > 0);
        NavigateAutomaticEditCommand = new RelayCommand(
            _ => NavigateToNextAutomaticEdit(),
            _ => !IsBusy && _session?.ProcessedMergeCount > 0);
        UndoCommand = new RelayCommand(_ => Undo(), _ => _undo.Count > 0 && !IsBusy);
        RedoCommand = new RelayCommand(_ => Redo(), _ => _redo.Count > 0 && !IsBusy);
        ExportDiagnosticsCommand = new RelayCommand(_ => ExportDiagnostics(), _ => _diagnosticLogger is not null && !IsBusy);
        CancelCommand = new RelayCommand(_ => CloseRequested?.Invoke());
        BrowseBaseCommand = new RelayCommand(_ => BasePath = BrowseWorkbook(BasePath));
        BrowseLocalCommand = new RelayCommand(_ => LocalPath = BrowseWorkbook(LocalPath));
        BrowseRemoteCommand = new RelayCommand(_ => RemotePath = BrowseWorkbook(RemotePath));
        BrowseOutputCommand = new RelayCommand(_ => OutputPath = BrowseOutput(OutputPath));
        BrowseTablesCommand = new RelayCommand(_ => TablesPath = BrowseCsv(TablesPath));
    }

    public ObservableCollection<ConflictItemViewModel> Conflicts { get; } = new();
    public ObservableCollection<SheetTabViewModel> SheetTabs { get; } = new();
    public ICollectionView ConflictsView { get; }
    public MergeGridTable? BaseGrid { get => _baseGrid; private set => SetField(ref _baseGrid, value); }
    public MergeGridTable? LocalGrid { get => _localGrid; private set => SetField(ref _localGrid, value); }
    public MergeGridTable? RemoteGrid { get => _remoteGrid; private set => SetField(ref _remoteGrid, value); }
    public MergeGridTable? MergedGrid { get => _mergedGrid; private set => SetField(ref _mergedGrid, value); }
    public IReadOnlyList<string> FilterOptions { get; } = new[] { "全部", "未解决", "已解决", "元数据复核", "内容冲突", "同键新增", "删除/修改" };
    public IReadOnlyList<RecalculationModeOption> RecalculationModes { get; }
    public ICommand LoadCommand { get; }
    public ICommand SaveCommand { get; }
    public ICommand ResolveCommand { get; }
    public ICommand PreviousCommand { get; }
    public ICommand NextCommand { get; }
    public ICommand NavigateAutomaticEditCommand { get; }
    public ICommand UndoCommand { get; }
    public ICommand RedoCommand { get; }
    public ICommand ExportDiagnosticsCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand BrowseBaseCommand { get; }
    public ICommand BrowseLocalCommand { get; }
    public ICommand BrowseRemoteCommand { get; }
    public ICommand BrowseOutputCommand { get; }
    public ICommand BrowseTablesCommand { get; }

    public string BasePath { get => _basePath; set => SetPath(ref _basePath, value); }
    public string LocalPath { get => _localPath; set => SetPath(ref _localPath, value); }
    public string RemotePath { get => _remotePath; set => SetPath(ref _remotePath, value); }
    public string OutputPath { get => _outputPath; set => SetPath(ref _outputPath, value); }
    public string RepositoryRoot { get => _repositoryRoot; set => SetPath(ref _repositoryRoot, value); }
    public string DataRoot { get => _dataRoot; set => SetPath(ref _dataRoot, value); }
    public string TablesPath { get => _tablesPath; set => SetPath(ref _tablesPath, value); }
    public RecalculationModeOption SelectedRecalculationMode
    {
        get => _selectedRecalculationMode;
        set => SetField(ref _selectedRecalculationMode, value);
    }

    public string SearchText
    {
        get => _searchText;
        set { if (SetField(ref _searchText, value)) ConflictsView.Refresh(); }
    }

    public string SelectedFilter
    {
        get => _selectedFilter;
        set { if (SetField(ref _selectedFilter, value)) ConflictsView.Refresh(); }
    }

    public ConflictItemViewModel? SelectedConflict
    {
        get => _selectedConflict;
        set
        {
            if (!SetField(ref _selectedConflict, value) || value is null)
                return;
            if (_settingSelectionFromGrid)
                return;
            _selectedConflictIds.Clear();
            _selectedConflictIds.Add(value.Id);
            NotifySelectionScopeChanged();
            ConflictNavigationRequested?.Invoke(value.Model.GridRowIndex, value.Model.GridColumnIndex);
        }
    }

    public SheetTabViewModel? SelectedSheet
    {
        get => _selectedSheet;
        set
        {
            if (!SetField(ref _selectedSheet, value) || value is null)
                return;
            ShowSheet(value);
        }
    }

    public event Action<int, int>? ConflictNavigationRequested;
    public event Action? ExternalMergeCompleted;
    public event Action? CloseRequested;
    public event Action<bool>? SavingStateChanged;

    public string StatusText { get => _statusText; private set => SetField(ref _statusText, value); }
    public bool IsBusy { get => _isBusy; private set { if (SetField(ref _isBusy, value)) RaiseCommandStates(); } }
    public bool IsSaving
    {
        get => _isSaving;
        private set
        {
            if (SetField(ref _isSaving, value))
                SavingStateChanged?.Invoke(value);
        }
    }
    public bool HasSession => _session is not null;
    public string LogicalTable => _session?.LogicalTable ?? "-";
    public string SheetName => SelectedSheet?.SheetName ?? "-";
    public string KeyName => SelectedSheet?.KeyName ?? "-";
    public int ConflictCount => _session?.Conflicts.Count ?? 0;
    public int RemainingCount => _session?.RemainingConflicts ?? 0;
    public int ResolvedCount => ConflictCount - RemainingCount;
    public int AutomaticMergeCount => _session?.AutomaticMergeCount ?? 0;
    public int ProcessedMergeCount => _session?.ProcessedMergeCount ?? 0;
    public string ProcessedMergeSummary
    {
        get
        {
            var count = ProcessedMergeCount;
            return _processedMergeNavigationIndex >= 0 && _processedMergeNavigationIndex < count
                ? $"已处理合并 {count} 格，当前 {_processedMergeNavigationIndex + 1}/{count}"
                : $"已处理合并 {count} 格";
        }
    }
    public int SelectedConflictCount => _selectedConflictIds.Count;
    public bool IsExternalMergeInvocation { get; private set; }
    public bool IsMergeCompleted { get; private set; }
    public int ExternalMergeExitCode { get; private set; } = ExitCodes.UnresolvedConflicts;
    public string SelectionText => SelectedConflictCount == 0
        ? "所选范围无冲突"
        : $"已选 {SelectedConflictCount} 个冲突";
    private bool HasRequiredInputs =>
        !string.IsNullOrWhiteSpace(BasePath) && !string.IsNullOrWhiteSpace(LocalPath) &&
        !string.IsNullOrWhiteSpace(RemotePath) && !string.IsNullOrWhiteSpace(OutputPath) &&
        !string.IsNullOrWhiteSpace(RepositoryRoot);

    public async Task LoadArgumentsAsync(IReadOnlyList<string> args)
    {
        IsExternalMergeInvocation = args.Count > 0;
        try
        {
            var options = EnsureDiagnosticLog(CommandLineParser.Parse(args));
            StartDiagnosticLog(options);
            PopulatePaths(options);
            await LoadAsync(options);
        }
        catch (Exception exception)
        {
            RecordFailure(exception);
            ShowError(exception);
        }
    }

    private async Task LoadFromFieldsAsync()
    {
        var options = EnsureDiagnosticLog(new MergeCommandOptions(
            BasePath, LocalPath, RemotePath, OutputPath, RepositoryRoot,
            NullIfBlank(DataRoot), NullIfBlank(TablesPath), null, "zh-CN",
            false, false, SelectedRecalculationMode.Value, null));
        StartDiagnosticLog(options);
        await LoadAsync(options);
    }

    private async Task LoadAsync(MergeCommandOptions options)
    {
        IsBusy = true;
        StatusText = "正在分析三个版本...";
        try
        {
            var session = await Task.Run(() => _coordinator.Prepare(options));
            _session = session;
            _processedMergeNavigationIndex = -1;
            _diagnosticLogger?.WritePrepared(session);
            _undo.Clear();
            _redo.Clear();
            _selectedConflictIds.Clear();
            SheetTabs.Clear();
            foreach (var sheet in session.Sheets)
                SheetTabs.Add(new SheetTabViewModel(sheet));
            SelectedSheet = SheetTabs.FirstOrDefault(sheet => sheet.HasUnresolvedMetadataChanges) ??
                            SheetTabs.FirstOrDefault(sheet => sheet.HasUnresolvedConflicts) ??
                            SheetTabs.FirstOrDefault();
            NotifySessionChanged();
            var analysisStatus = session.MetadataChangeCount > 0
                ? $"检测到 {session.MetadataChangeCount} 处 Luban 元数据变更，请逐项确认后保存"
                : session.RequiresStructuralChangeConfirmation
                ? $"检测到 {session.StructuralChanges.Count} 项列结构变化，保存前需要二次确认"
                : session.Conflicts.Count == 0
                ? $"分析完成：{session.Sheets.Count} 个工作表，{session.AutomaticMergeCount} 项自动合并结果，可直接保存"
                : $"分析完成：{session.Sheets.Count} 个工作表，{session.Conflicts.Count} 个冲突待处理";
            var configurationStatus = new List<string>();
            if (session.IgnoredFields.Count > 0)
                configurationStatus.Add($"忽略字段保留 LOCAL：{string.Join("、", session.IgnoredFields)}");
            if (session.LogicalTableUniquenessValidated)
                configurationStatus.Add("已检查全逻辑表唯一性");
            StatusText = configurationStatus.Count == 0
                ? analysisStatus
                : $"{analysisStatus}；{string.Join("；", configurationStatus)}";
        }
        catch (Exception exception)
        {
            RecordFailure(exception);
            ShowError(exception);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task SaveAsync()
    {
        if (_session is null)
            return;
        if (_session.RequiresStructuralChangeConfirmation &&
            !_confirmStructuralChanges(_session.StructuralChanges))
        {
            StatusText = "已取消保存，MERGED 未被修改";
            return;
        }
        IsBusy = true;
        IsSaving = true;
        StatusText = "正在原子保存并验证工作簿...";
        Exception? saveFailure = null;
        var externalMergeCompleted = false;
        try
        {
            var result = await Task.Run(_session.Save);
            GitStageResult? gitStageResult = null;
            if (IsExternalMergeInvocation)
            {
                try
                {
                    gitStageResult = await Task.Run(() => _gitStager.Stage(RepositoryRoot, result.OutputPath));
                }
                catch (GitStagingException exception)
                {
                    throw new WorkbookWriteException(
                        "MERGED 已保存，但无法在 Git 中标记 resolved 并加入 staged。请保留当前文件并检查 Git/Fork 配置。",
                        exception);
                }
            }
            _diagnosticLogger?.WriteSaved(_session, result, gitStageResult is not null);
            var validationText = result.ProjectValidationCompleted ? "，项目快速校验通过" : string.Empty;
            if (result.FullExportValidationCompleted)
                validationText += "，隔离完整导出通过";
            if (gitStageResult is not null)
                validationText += "，已在 Git 中标记 resolved 并加入 staged";
            StatusText = result.RecalculationStatus switch
            {
                LubanExcelMerge.OpenXml.WorkbookRecalculationStatus.Completed =>
                    $"保存完成并已由 {result.RecalculationProvider} 完整重算{validationText}：{result.OutputPath}",
                LubanExcelMerge.OpenXml.WorkbookRecalculationStatus.DeferredAfterRecalculationFailure =>
                    $"保存完成：{result.RecalculationProvider} 重算未完成，已安全回退并标记下次打开完整重算{validationText}：{result.OutputPath}",
                LubanExcelMerge.OpenXml.WorkbookRecalculationStatus.SourceCachePreservedUnverified =>
                    $"保存完成：已保留公式缓存并标记下次完整重算（缓存未验证）{validationText}：{result.OutputPath}",
                _ => $"保存完成{validationText}：{result.OutputPath}"
            };
            if (IsExternalMergeInvocation)
            {
                IsMergeCompleted = true;
                ExternalMergeExitCode = ExitCodes.Success;
                externalMergeCompleted = true;
            }
        }
        catch (Exception exception)
        {
            RecordFailure(exception);
            saveFailure = exception;
        }
        finally
        {
            IsSaving = false;
            IsBusy = false;
        }

        if (saveFailure is not null)
        {
            ShowError(saveFailure);
            return;
        }
        if (externalMergeCompleted)
            ExternalMergeCompleted?.Invoke();
    }

    private void ResolveSelected(object? parameter)
    {
        if (parameter is not MergeChoice choice)
            return;
        var selectedItems = Conflicts.Where(item => _selectedConflictIds.Contains(item.Id)).ToArray();
        var changes = selectedItems
            .Where(item => item.Model.SelectedChoice != choice)
            .Select(item => new ResolutionChange(item, item.Model.SelectedChoice, choice))
            .ToArray();
        if (changes.Length == 0)
            return;

        foreach (var change in changes)
            SetResolution(change.Item, change.NewChoice);
        RefreshAfterResolutions();
        _undo.Push(new ResolutionBatch(changes));
        _redo.Clear();
        StatusText = RemainingCount == 0
            ? $"已批量解决 {changes.Length} 个冲突，所有冲突均已解决"
            : $"已批量解决 {changes.Length} 个冲突，剩余 {RemainingCount} 个";
        RaiseCommandStates();
    }

    private void Undo()
    {
        var batch = _undo.Pop();
        foreach (var change in batch.Changes)
            SetResolution(change.Item, change.OldChoice);
        SelectBatchSheet(batch);
        RefreshAfterResolutions();
        _redo.Push(batch);
        StatusText = $"已撤销 {batch.Changes.Count} 个冲突的选择";
        RaiseCommandStates();
    }

    private void Redo()
    {
        var batch = _redo.Pop();
        foreach (var change in batch.Changes)
            SetResolution(change.Item, change.NewChoice);
        SelectBatchSheet(batch);
        RefreshAfterResolutions();
        _undo.Push(batch);
        StatusText = $"已重做 {batch.Changes.Count} 个冲突的选择";
        RaiseCommandStates();
    }

    private static void SetResolution(ConflictItemViewModel item, MergeChoice? choice)
    {
        if (choice is { } selected)
            item.Model.Resolve(selected);
        else
            item.Model.ClearResolution();
        item.Refresh();
    }

    private void SelectBatchSheet(ResolutionBatch batch)
    {
        var first = batch.Changes.FirstOrDefault()?.Item;
        if (first is null)
            return;
        var sheet = SheetTabs.FirstOrDefault(candidate => candidate.Conflicts.Contains(first));
        if (sheet is not null && !ReferenceEquals(sheet, SelectedSheet))
            SelectedSheet = sheet;
        SelectedConflict = first;
    }

    private void RefreshAfterResolutions()
    {
        _processedMergeNavigationIndex = -1;
        ConflictsView.Refresh();
        RefreshAllGrids();
        foreach (var sheet in SheetTabs)
            sheet.Refresh();
        NotifySessionChanged();
    }

    private void MoveSelection(int offset)
    {
        var allConflicts = SheetTabs
            .SelectMany(sheet => sheet.Conflicts.Select(conflict => (Sheet: sheet, Conflict: conflict)))
            .ToArray();
        if (allConflicts.Length == 0)
            return;
        var current = SelectedConflict is null || SelectedSheet is null
            ? -1
            : Array.FindIndex(allConflicts, item =>
                ReferenceEquals(item.Sheet, SelectedSheet) && ReferenceEquals(item.Conflict, SelectedConflict));
        var next = current < 0 ? 0 : (current + offset + allConflicts.Length) % allConflicts.Length;
        var target = allConflicts[next];
        if (!ReferenceEquals(SelectedSheet, target.Sheet))
            SelectedSheet = target.Sheet;
        SelectedConflict = target.Conflict;
    }

    private void NavigateToNextAutomaticEdit()
    {
        var targets = SheetTabs
            .SelectMany(sheet => sheet.Model.ProcessedMergeLocations.Select(location =>
                (Sheet: sheet, Location: location)))
            .ToArray();
        if (targets.Length == 0)
        {
            _processedMergeNavigationIndex = -1;
            OnPropertyChanged(nameof(ProcessedMergeSummary));
            return;
        }

        _processedMergeNavigationIndex = (_processedMergeNavigationIndex + 1) % targets.Length;
        var target = targets[_processedMergeNavigationIndex];
        if (!ReferenceEquals(SelectedSheet, target.Sheet))
            SelectedSheet = target.Sheet;

        SelectedConflict = null;
        _selectedConflictIds.Clear();
        NotifySelectionScopeChanged();
        ConflictNavigationRequested?.Invoke(target.Location.RowIndex, target.Location.ColumnIndex);
        StatusText = $"已定位处理结果 {_processedMergeNavigationIndex + 1}/{targets.Length}：" +
                      $"{target.Sheet.SheetName}!{target.Location.DisplayLocation}";
        OnPropertyChanged(nameof(ProcessedMergeSummary));
    }

    public void SelectConflict(string conflictId)
    {
        var item = Conflicts.FirstOrDefault(conflict => conflict.Id == conflictId);
        if (item is not null)
            SelectedConflict = item;
    }

    public void SelectConflicts(IReadOnlyList<string> conflictIds)
    {
        _selectedConflictIds.Clear();
        foreach (var id in conflictIds.Where(id => Conflicts.Any(conflict => conflict.Id == id)))
            _selectedConflictIds.Add(id);

        var primary = Conflicts.FirstOrDefault(conflict => _selectedConflictIds.Contains(conflict.Id));
        if (primary is not null)
        {
            _settingSelectionFromGrid = true;
            try
            {
                SelectedConflict = primary;
            }
            finally
            {
                _settingSelectionFromGrid = false;
            }
        }
        NotifySelectionScopeChanged();
    }

    private void NotifySelectionScopeChanged()
    {
        OnPropertyChanged(nameof(SelectedConflictCount));
        OnPropertyChanged(nameof(SelectionText));
        RaiseCommandStates();
    }

    private void RefreshAllGrids()
    {
        if (SelectedSheet is null)
            return;
        BaseGrid = SelectedSheet.Model.Comparison.CreateTable(MergeGridSide.Base);
        LocalGrid = SelectedSheet.Model.Comparison.CreateTable(MergeGridSide.Local);
        RemoteGrid = SelectedSheet.Model.Comparison.CreateTable(MergeGridSide.Remote);
        MergedGrid = SelectedSheet.Model.Comparison.CreateTable(MergeGridSide.Merged);
    }

    private void ShowSheet(SheetTabViewModel sheet)
    {
        _selectedConflictIds.Clear();
        Conflicts.Clear();
        foreach (var conflict in sheet.Conflicts)
            Conflicts.Add(conflict);
        RefreshAllGrids();
        SelectedConflict = Conflicts.FirstOrDefault(conflict => conflict.IsMetadataChange && !conflict.IsResolved) ??
                           Conflicts.FirstOrDefault(conflict => !conflict.IsResolved) ??
                           Conflicts.FirstOrDefault();
        ConflictsView.Refresh();
        OnPropertyChanged(nameof(SheetName));
        OnPropertyChanged(nameof(KeyName));
        NotifySelectionScopeChanged();
    }

    private bool FilterConflict(object value)
    {
        if (value is not ConflictItemViewModel item)
            return false;
        var matchesFilter = SelectedFilter switch
        {
            "未解决" => !item.IsResolved,
            "已解决" => item.IsResolved,
            "元数据复核" => item.IsMetadataChange,
            "内容冲突" => item.KindText == "内容冲突",
            "同键新增" => item.KindText == "同键新增",
            "删除/修改" => item.KindText == "删除/修改",
            _ => true
        };
        if (!matchesFilter || string.IsNullOrWhiteSpace(SearchText))
            return matchesFilter;
        return new[] { item.RecordKey, item.FieldName, item.BaseValue, item.LocalValue, item.RemoteValue }
            .Any(text => text.Contains(SearchText, StringComparison.OrdinalIgnoreCase));
    }

    private void PopulatePaths(MergeCommandOptions options)
    {
        BasePath = options.BasePath;
        LocalPath = options.LocalPath;
        RemotePath = options.RemotePath;
        OutputPath = options.OutputPath;
        RepositoryRoot = options.RepositoryRoot;
        DataRoot = options.DataRoot ?? string.Empty;
        TablesPath = options.TablesPath ?? string.Empty;
        if (options.RecalculateWithExcel is not null)
            SelectedRecalculationMode = RecalculationModes.First(mode => mode.Value == options.RecalculateWithExcel);
    }

    private static MergeCommandOptions EnsureDiagnosticLog(MergeCommandOptions options) =>
        string.IsNullOrWhiteSpace(options.LogPath)
            ? options with { LogPath = MergeDiagnosticLogger.CreateDefaultPath() }
            : options;

    private void StartDiagnosticLog(MergeCommandOptions options)
    {
        _diagnosticLogger = MergeDiagnosticLogger.Create(options.LogPath);
        _diagnosticLogger?.WriteStarted(options);
    }

    private void RecordFailure(Exception exception)
    {
        var exitCode = ExitCodes.ForException(exception);
        if (IsExternalMergeInvocation)
            ExternalMergeExitCode = exitCode;
        _diagnosticLogger?.WriteException(exception, exitCode);
    }

    private void ExportDiagnostics()
    {
        if (_diagnosticLogger is null)
            return;
        var dialog = new SaveFileDialog
        {
            Filter = "ZIP 诊断包 (*.zip)|*.zip",
            FileName = $"LubanExcelMerge-diagnostics-{DateTime.Now:yyyyMMdd-HHmmss}.zip"
        };
        if (dialog.ShowDialog() != true)
            return;
        try
        {
            var result = DiagnosticPackageExporter.Export(_diagnosticLogger.Path, dialog.FileName);
            StatusText = $"诊断包已导出（不含工作簿）：{result.OutputPath}";
        }
        catch (Exception exception)
        {
            RecordFailure(exception);
            ShowError(exception);
        }
    }

    private void NotifySessionChanged()
    {
        OnPropertyChanged(nameof(HasSession));
        OnPropertyChanged(nameof(LogicalTable));
        OnPropertyChanged(nameof(SheetName));
        OnPropertyChanged(nameof(KeyName));
        OnPropertyChanged(nameof(ConflictCount));
        OnPropertyChanged(nameof(RemainingCount));
        OnPropertyChanged(nameof(ResolvedCount));
        OnPropertyChanged(nameof(AutomaticMergeCount));
        OnPropertyChanged(nameof(ProcessedMergeCount));
        OnPropertyChanged(nameof(ProcessedMergeSummary));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        (LoadCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (SaveCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (ResolveCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (PreviousCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (NextCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (NavigateAutomaticEditCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (UndoCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (RedoCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (ExportDiagnosticsCommand as RelayCommand)?.RaiseCanExecuteChanged();
    }

    private void ShowError(Exception exception)
    {
        var gitException = FindInnerException<GitStagingException>(exception);
        var detail = gitException?.Message ?? FindInnermostException(exception)?.Message;
        var detailLabel = gitException is null ? "详细原因" : "Git 详细原因";
        var message = string.IsNullOrWhiteSpace(detail)
            ? exception.Message
            : $"{exception.Message}{Environment.NewLine}{Environment.NewLine}{detailLabel}：{detail}";
        StatusText = message;
        MessageBox.Show(message, "Luban Excel 合并", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static bool ConfirmStructuralChanges(IReadOnlyList<string> changes)
    {
        var displayedChanges = changes.Take(12).Select(change => $"• {change}").ToList();
        if (changes.Count > displayedChanges.Count)
            displayedChanges.Add($"• 另有 {changes.Count - displayedChanges.Count} 项未展开");
        var message =
            "本次保存将删除、修改或移动既有字段列。列结构变化可能影响公式引用和下游 Luban 配置。" +
            Environment.NewLine + Environment.NewLine +
            string.Join(Environment.NewLine, displayedChanges) +
            Environment.NewLine + Environment.NewLine +
            "是否确认继续保存并执行后续校验与 Git staging？";
        return MessageBox.Show(
            message,
            "确认保存列结构变更",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No) == MessageBoxResult.Yes;
    }

    private static TException? FindInnerException<TException>(Exception exception)
        where TException : Exception
    {
        for (var current = exception; current is not null; current = current.InnerException!)
        {
            if (current is TException match)
                return match;
        }
        return null;
    }

    private static Exception? FindInnermostException(Exception exception)
    {
        var current = exception.InnerException;
        if (current is null)
            return null;
        while (current.InnerException is not null)
            current = current.InnerException;
        return current;
    }

    private void SetPath(ref string field, string value, [CallerMemberName] string? name = null)
    {
        if (string.Equals(field, value, StringComparison.Ordinal))
            return;
        field = value;
        OnPropertyChanged(name);
        RaiseCommandStates();
    }

    private static string BrowseWorkbook(string current) => BrowseOpen(current, "Excel 工作簿 (*.xlsx)|*.xlsx");
    private static string BrowseCsv(string current) => BrowseOpen(current, "CSV 文件 (*.csv)|*.csv");
    private static string BrowseOpen(string current, string filter)
    {
        var dialog = new OpenFileDialog { Filter = filter, FileName = current };
        return dialog.ShowDialog() == true ? dialog.FileName : current;
    }

    private static string BrowseOutput(string current)
    {
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", FileName = current };
        return dialog.ShowDialog() == true ? dialog.FileName : current;
    }

    private static string? NullIfBlank(string value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;
        field = value;
        OnPropertyChanged(name);
        return true;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    private sealed record ResolutionChange(
        ConflictItemViewModel Item,
        MergeChoice? OldChoice,
        MergeChoice NewChoice);

    private sealed record ResolutionBatch(IReadOnlyList<ResolutionChange> Changes);
}
