using LubanExcelMerge.Core;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Cli;

public sealed record MergePreparationTimings(
    long BaseWorkbookReadMilliseconds,
    long LocalWorkbookReadMilliseconds,
    long RemoteWorkbookReadMilliseconds,
    long SheetPreparationMilliseconds,
    long TotalMilliseconds)
{
    public long WorkbookReadMilliseconds =>
        BaseWorkbookReadMilliseconds + LocalWorkbookReadMilliseconds + RemoteWorkbookReadMilliseconds;
}

public enum MergeChoice
{
    Base,
    Local,
    Remote
}

public sealed class ResolvableMergeConflict
{
    private readonly IReadOnlyDictionary<MergeChoice, IReadOnlyList<WorkbookEdit>> _choiceEdits;

    internal ResolvableMergeConflict(
        string id,
        MergeConflict conflict,
        int? rowNumber,
        string baseValue,
        string localValue,
        string remoteValue,
        IReadOnlyDictionary<MergeChoice, IReadOnlyList<WorkbookEdit>> choiceEdits)
    {
        Id = id;
        Conflict = conflict;
        RowNumber = rowNumber;
        BaseValue = baseValue;
        LocalValue = localValue;
        RemoteValue = remoteValue;
        _choiceEdits = choiceEdits;
    }

    public string Id { get; }
    public MergeConflict Conflict { get; }
    public int? RowNumber { get; }
    public int GridRowIndex { get; private set; } = -1;
    public int GridColumnIndex { get; private set; } = -1;
    public string BaseValue { get; }
    public string LocalValue { get; }
    public string RemoteValue { get; }
    public MergeChoice? SelectedChoice { get; private set; }
    public bool IsResolved => SelectedChoice is not null;

    public void Resolve(MergeChoice choice)
    {
        if (!_choiceEdits.ContainsKey(choice))
            throw new ArgumentOutOfRangeException(nameof(choice));
        SelectedChoice = choice;
    }

    public void ClearResolution() => SelectedChoice = null;

    internal void SetGridLocation(int rowIndex, int columnIndex)
    {
        GridRowIndex = rowIndex;
        GridColumnIndex = columnIndex;
    }

    internal IReadOnlyList<WorkbookEdit> GetSelectedEdits() =>
        SelectedChoice is { } choice
            ? _choiceEdits[choice]
            : throw new InvalidOperationException($"冲突 {Id} 尚未解决。");
}

public sealed class PreparedSheetMerge
{
    private IReadOnlyList<MergeGridLocation>? _automaticEditLocations;
    private IReadOnlyList<MergeGridLocation>? _automaticMergeLocations;
    private readonly Func<IReadOnlyList<WorkbookEdit>>? _finalEditFactory;

    internal PreparedSheetMerge(
        string sheetName,
        string keyName,
        IReadOnlyList<WorkbookEdit> automaticEdits,
        IReadOnlyList<ResolvableMergeConflict> conflicts,
        MergeComparison comparison,
        int changedCells,
        int addedRecords,
        int deletedRecords,
        int metadataChangeCount,
        IReadOnlyList<string>? structuralChanges = null,
        Func<IReadOnlyList<WorkbookEdit>>? finalEditFactory = null)
    {
        SheetName = sheetName;
        KeyName = keyName;
        AutomaticEdits = automaticEdits;
        Conflicts = conflicts;
        Comparison = comparison;
        ChangedCells = changedCells;
        AddedRecords = addedRecords;
        DeletedRecords = deletedRecords;
        MetadataChangeCount = metadataChangeCount;
        StructuralChanges = structuralChanges ?? Array.Empty<string>();
        _finalEditFactory = finalEditFactory;
    }

    public string SheetName { get; }
    public string KeyName { get; }
    public IReadOnlyList<ResolvableMergeConflict> Conflicts { get; }
    public MergeComparison Comparison { get; }
    public int ChangedCells { get; }
    public int AddedRecords { get; }
    public int DeletedRecords { get; }
    public int MetadataChangeCount { get; }
    public IReadOnlyList<string> StructuralChanges { get; }
    public bool RequiresStructuralChangeConfirmation => StructuralChanges.Count > 0;
    public int AutomaticEditCount => AutomaticEdits.Count;
    public int AutomaticMergeCount => AutomaticMergeLocations.Count;
    public IReadOnlyList<MergeGridLocation> AutomaticMergeLocations =>
        _automaticMergeLocations ??= Comparison.FindAutomaticMergeLocations();
    public int ProcessedMergeCount => ProcessedMergeLocations.Count;
    public IReadOnlyList<MergeGridLocation> ProcessedMergeLocations =>
        Comparison.FindProcessedMergeLocations();
    public IReadOnlyList<MergeGridLocation> AutomaticEditLocations
    {
        get
        {
            if (_automaticEditLocations is not null)
                return _automaticEditLocations;
            var locations = new List<MergeGridLocation>(AutomaticEdits.Count);
            foreach (var edit in AutomaticEdits)
            {
                var location = Comparison.FindAutomaticEditLocation(edit);
                if (location is not null)
                    locations.Add(location);
            }

            _automaticEditLocations = locations;
            return _automaticEditLocations;
        }
    }
    public MergeGridLocation? FirstAutomaticEditLocation => AutomaticEditLocations.FirstOrDefault();
    public int RemainingConflicts => Conflicts.Count(conflict => !conflict.IsResolved);
    public bool CanSave => RemainingConflicts == 0;
    internal IReadOnlyList<WorkbookEdit> AutomaticEdits { get; }

    internal IEnumerable<WorkbookEdit> GetSelectedEdits() =>
        _finalEditFactory?.Invoke() ??
        AutomaticEdits.Concat(Conflicts.SelectMany(conflict => conflict.GetSelectedEdits()));
}

public sealed class PreparedMergeSession
{
    private readonly AtomicWorkbookSaver _saver;
    private readonly int _localFormulaCount;
    private readonly WorkbookRecalculationMode _recalculationMode;
    private readonly string _repositoryRoot;
    private readonly bool _projectValidationEnabled;
    private readonly string? _projectValidationCommand;
    private readonly IProjectValidator _projectValidator;
    private readonly bool _fullExportValidationEnabled;
    private readonly string? _fullExportValidationCommand;
    private readonly IFullExportValidator _fullExportValidator;

    internal PreparedMergeSession(
        string logicalTable,
        string localPath,
        string outputPath,
        IReadOnlyList<string> ignoredFields,
        bool logicalTableUniquenessValidated,
        IReadOnlyList<PreparedSheetMerge> sheets,
        int localFormulaCount,
        WorkbookRecalculationMode recalculationMode,
        AtomicWorkbookSaver saver,
        string repositoryRoot,
        bool projectValidationEnabled,
        string? projectValidationCommand,
        IProjectValidator projectValidator,
        bool fullExportValidationEnabled,
        string? fullExportValidationCommand,
        IFullExportValidator fullExportValidator,
        MergePreparationTimings preparationTimings)
    {
        if (sheets.Count == 0)
            throw new ArgumentException("合并会话必须包含至少一个工作表。", nameof(sheets));
        LogicalTable = logicalTable;
        LocalPath = localPath;
        OutputPath = outputPath;
        IgnoredFields = ignoredFields;
        LogicalTableUniquenessValidated = logicalTableUniquenessValidated;
        Sheets = sheets;
        Conflicts = sheets.SelectMany(sheet => sheet.Conflicts).ToArray();
        _localFormulaCount = localFormulaCount;
        _recalculationMode = recalculationMode;
        _saver = saver;
        _repositoryRoot = repositoryRoot;
        _projectValidationEnabled = projectValidationEnabled;
        _projectValidationCommand = projectValidationCommand;
        _projectValidator = projectValidator;
        _fullExportValidationEnabled = fullExportValidationEnabled;
        _fullExportValidationCommand = fullExportValidationCommand;
        _fullExportValidator = fullExportValidator;
        PreparationTimings = preparationTimings;
    }

    public string LogicalTable { get; }
    public string SheetName => Sheets.Count == 1
        ? Sheets[0].SheetName
        : string.Join(", ", Sheets.Select(sheet => sheet.SheetName));
    public string KeyName => Sheets.Select(sheet => sheet.KeyName).Distinct(StringComparer.Ordinal).ToArray() switch
    {
    [var key] => key,
        _ => "按工作表"
    };
    public string LocalPath { get; }
    public string OutputPath { get; }
    public IReadOnlyList<string> IgnoredFields { get; }
    public bool LogicalTableUniquenessValidated { get; }
    public MergePreparationTimings PreparationTimings { get; }
    public IReadOnlyList<PreparedSheetMerge> Sheets { get; }
    public IReadOnlyList<ResolvableMergeConflict> Conflicts { get; }
    public MergeComparison Comparison => Sheets[0].Comparison;
    public int ChangedCells => Sheets.Sum(sheet => sheet.ChangedCells);
    public int AddedRecords => Sheets.Sum(sheet => sheet.AddedRecords);
    public int DeletedRecords => Sheets.Sum(sheet => sheet.DeletedRecords);
    public int AutomaticEditCount => Sheets.Sum(sheet => sheet.AutomaticEditCount);
    public int AutomaticMergeCount => Sheets.Sum(sheet => sheet.AutomaticMergeCount);
    public int ProcessedMergeCount => Sheets.Sum(sheet => sheet.ProcessedMergeCount);
    public int MetadataChangeCount => Sheets.Sum(sheet => sheet.MetadataChangeCount);
    public IReadOnlyList<string> StructuralChanges => Sheets
        .SelectMany(sheet => sheet.StructuralChanges.Select(change => $"{sheet.SheetName}：{change}"))
        .ToArray();
    public bool RequiresStructuralChangeConfirmation =>
        Sheets.Any(sheet => sheet.RequiresStructuralChangeConfirmation);
    public int RemainingConflicts => Conflicts.Count(conflict => !conflict.IsResolved);
    public bool CanSave => RemainingConflicts == 0;

    public WorkbookSaveResult Save()
    {
        if (!CanSave)
            throw new InvalidOperationException($"仍有 {RemainingConflicts} 个冲突未解决，不能保存。");

        var mergeEdits = Sheets.SelectMany(sheet => sheet.GetSelectedEdits()).ToArray();
        var formulaMayBeAffected = mergeEdits.Length > 0 &&
            (_localFormulaCount > 0 || mergeEdits.Any(ContainsFormula));
        var cleanupEdits = Sheets.Select(sheet => (WorkbookEdit)new CleanupEmptyFormattingEdit(sheet.SheetName));
        var edits = cleanupEdits
            .Concat(mergeEdits)
            .Concat(cleanupEdits)
            .ToArray();
        try
        {
            using var outputLease = MergeOutputLease.Acquire(OutputPath);
            MergeOutputRecovery.RecoverPending(OutputPath);
            return _projectValidationEnabled || _fullExportValidationEnabled
                ? SaveAndValidate(edits, formulaMayBeAffected)
                : SaveWorkbook(edits, formulaMayBeAffected);
        }
        catch (WorkbookRecalculationUnavailableException exception)
        {
            throw new UnsafeWorkbookException(exception.Message, exception);
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or InvalidOperationException)
        {
            throw new WorkbookWriteException("写入或重新打开 MERGED 验证失败。", exception);
        }
    }

    private WorkbookSaveResult SaveAndValidate(
        IReadOnlyList<WorkbookEdit> edits,
        bool formulaMayBeAffected)
    {
        if (_projectValidationEnabled && string.IsNullOrWhiteSpace(_projectValidationCommand))
            throw new ProjectValidationException("已启用项目快速校验，但没有可执行的校验命令。");
        if (_fullExportValidationEnabled && string.IsNullOrWhiteSpace(_fullExportValidationCommand))
            throw new ProjectValidationException("已启用完整导出校验，但没有可执行的校验命令。");

        var outputExisted = File.Exists(OutputPath);
        var rollbackMarker = MergeOutputRecovery.CreateRollbackMarker(OutputPath, outputExisted);

        try
        {
            var result = SaveWorkbook(edits, formulaMayBeAffected);
            try
            {
                if (_projectValidationEnabled)
                {
                    _projectValidator.Validate(
                        _projectValidationCommand!,
                        _repositoryRoot,
                        TimeSpan.FromMinutes(10));
                }
                if (_fullExportValidationEnabled)
                {
                    _fullExportValidator.Validate(
                        _fullExportValidationCommand!,
                        _repositoryRoot,
                        TimeSpan.FromMinutes(30));
                }
            }
            catch (Exception exception)
            {
                if (exception is ProjectValidationException)
                    throw;
                throw new ProjectValidationException("运行项目快速校验失败。", exception);
            }

            var completedResult = result with
            {
                ProjectValidationCompleted = _projectValidationEnabled,
                FullExportValidationCompleted = _fullExportValidationEnabled
            };
            MergeOutputRecovery.Commit(rollbackMarker);
            return completedResult;
        }
        catch
        {
            if (File.Exists(rollbackMarker))
                MergeOutputRecovery.Restore(OutputPath, outputExisted, rollbackMarker);
            throw;
        }
    }

    private WorkbookSaveResult SaveWorkbook(
        IReadOnlyList<WorkbookEdit> edits,
        bool formulaMayBeAffected) =>
        _saver.Save(
            LocalPath,
            OutputPath,
            edits,
            new WorkbookSaveOptions(_recalculationMode, formulaMayBeAffected));

    private static bool ContainsFormula(WorkbookEdit edit) => edit switch
    {
        SetCellEdit setCell => setCell.Payload.Kind == CellValueKind.Formula,
        AppendRowEdit appendRow => appendRow.Cells.Any(cell => cell.Payload.Kind == CellValueKind.Formula),
        _ => false
    };
}
