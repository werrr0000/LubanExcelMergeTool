using LubanExcelMerge.Core;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Cli;

public enum MergeGridSide
{
    Base,
    Local,
    Remote,
    Merged
}

public enum MergeGridCellState
{
    Normal,
    Modified,
    Conflict,
    Added,
    Deleted,
    Metadata
}

public sealed record MergeGridCell(
    string DisplayValue,
    MergeGridCellState State,
    string? ConflictId = null);

public sealed record MergeGridRow(
    string RowHeader,
    string RecordKey,
    IReadOnlyList<MergeGridCell> Cells);

public sealed record MergeGridTable(
    string Title,
    IReadOnlyList<string> ColumnHeaders,
    IReadOnlyList<MergeGridRow> Rows);

public sealed record MergeGridLocation(int RowIndex, int ColumnIndex, string DisplayLocation);

public sealed class MergeComparison
{
    private readonly IReadOnlyList<ComparisonRowPlan> _rows;
    private readonly IReadOnlyDictionary<string, ResolvableMergeConflict> _conflicts;
    private readonly IReadOnlySet<int> _ignoredColumns;
    private readonly IReadOnlyDictionary<int, int> _localRowIndices;
    private readonly IReadOnlyDictionary<int, int> _remoteRowIndices;
    private readonly Func<int, bool> _mergedColumnIncluded;
    private readonly Func<int, string, CellPayload, CellPayload, CellPayload, CellPayload?> _structuralCellMerge;
    private MergeGridTable? _baseTable;
    private MergeGridTable? _localTable;
    private MergeGridTable? _remoteTable;

    internal MergeComparison(
        IReadOnlyList<string> columnHeaders,
        IReadOnlyList<ComparisonRowPlan> rows,
        IReadOnlyList<ResolvableMergeConflict> conflicts,
        IReadOnlySet<int> ignoredColumns,
        Func<int, bool>? mergedColumnIncluded = null,
        Func<int, string, CellPayload, CellPayload, CellPayload, CellPayload?>? structuralCellMerge = null)
    {
        ColumnHeaders = columnHeaders;
        _rows = rows;
        _conflicts = conflicts.ToDictionary(conflict => conflict.Id, StringComparer.Ordinal);
        _ignoredColumns = ignoredColumns;
        _mergedColumnIncluded = mergedColumnIncluded ?? (_ => true);
        _structuralCellMerge = structuralCellMerge ?? ((_, _, _, _, _) => null);
        _localRowIndices = rows
            .Select((row, index) => (row.LocalRowNumber, Index: index))
            .Where(item => item.LocalRowNumber is not null)
            .ToDictionary(item => item.LocalRowNumber!.Value, item => item.Index);
        _remoteRowIndices = rows
            .Select((row, index) => (row.RemoteRowNumber, Index: index))
            .Where(item => item.RemoteRowNumber is not null)
            .ToDictionary(item => item.RemoteRowNumber!.Value, item => item.Index);
    }

    public IReadOnlyList<string> ColumnHeaders { get; }

    public MergeGridTable CreateTable(MergeGridSide side) => side switch
    {
        MergeGridSide.Base => _baseTable ??= CreateTableCore(side),
        MergeGridSide.Local => _localTable ??= CreateTableCore(side),
        MergeGridSide.Remote => _remoteTable ??= CreateTableCore(side),
        MergeGridSide.Merged => CreateTableCore(side),
        _ => throw new ArgumentOutOfRangeException(nameof(side))
    };

    private MergeGridTable CreateTableCore(MergeGridSide side) => new(
        side switch
        {
            MergeGridSide.Base => "BASE",
            MergeGridSide.Local => "LOCAL",
            MergeGridSide.Remote => "REMOTE",
            MergeGridSide.Merged => "MERGED",
            _ => throw new ArgumentOutOfRangeException(nameof(side))
        },
        ColumnHeaders,
        new LazyReadOnlyList<MergeGridRow>(_rows.Count, rowIndex => CreateRow(_rows[rowIndex], side)));

    public MergeGridLocation? FindAutomaticEditLocation(WorkbookEdit edit)
    {
        return edit switch
        {
            SetCellEdit setCell => FindSetCellLocation(setCell),
            DeleteRowEdit deleteRow => FindDeleteRowLocation(deleteRow),
            AppendRowEdit appendRow => FindAppendRowLocation(appendRow),
            _ => null
        };
    }

    private MergeGridLocation? FindSetCellLocation(SetCellEdit edit)
    {
        var (rowNumber, columnIndex) = CellReference.Parse(edit.Address);
        return _localRowIndices.TryGetValue(rowNumber, out var rowIndex)
            ? new MergeGridLocation(rowIndex, columnIndex, edit.Address)
            : null;
    }

    private MergeGridLocation? FindDeleteRowLocation(DeleteRowEdit edit)
    {
        if (_localRowIndices.TryGetValue(edit.RowNumber, out var rowIndex))
        {
            var row = _rows[rowIndex];
            if (!row.IsStructure)
            {
                return new MergeGridLocation(
                    rowIndex,
                    FindFirstPopulatedColumn(row.LocalCells),
                    $"第 {edit.RowNumber} 行（删除）");
            }
        }

        return null;
    }

    private MergeGridLocation? FindAppendRowLocation(AppendRowEdit edit)
    {
        if (edit.SourceRowNumber is int sourceRowNumber &&
            _remoteRowIndices.TryGetValue(sourceRowNumber, out var indexedRow) &&
            IsMatchingRemoteAddition(_rows[indexedRow], edit))
        {
            return CreateAppendLocation(indexedRow, _rows[indexedRow]);
        }

        for (var rowIndex = 0; rowIndex < _rows.Count; rowIndex++)
        {
            var row = _rows[rowIndex];
            if (IsMatchingRemoteAddition(row, edit))
                return CreateAppendLocation(rowIndex, row);
        }

        return null;
    }

    private static bool IsMatchingRemoteAddition(ComparisonRowPlan row, AppendRowEdit edit) =>
        !row.IsStructure &&
        row.BaseCells is null &&
        row.LocalCells is null &&
        row.RemoteCells is not null &&
        edit.Cells.All(cell =>
            cell.ColumnIndex < row.RemoteCells.Length &&
            row.RemoteCells[cell.ColumnIndex].ContentEquals(cell.Payload));

    private static MergeGridLocation CreateAppendLocation(int rowIndex, ComparisonRowPlan row)
    {
        var displayLocation = row.RemoteRowNumber is int remoteRow
            ? $"第 {remoteRow} 行（新增）"
            : "新增行";
        return new MergeGridLocation(
            rowIndex,
            FindFirstPopulatedColumn(row.RemoteCells),
            displayLocation);
    }

    private static int FindFirstPopulatedColumn(IReadOnlyList<CellPayload?>? cells)
    {
        if (cells is null)
            return 0;

        for (var columnIndex = 0; columnIndex < cells.Count; columnIndex++)
        {
            if (cells[columnIndex] is { Kind: not CellValueKind.Blank })
                return columnIndex;
        }

        return 0;
    }

    private MergeGridRow CreateRow(ComparisonRowPlan row, MergeGridSide side)
    {
        var payloads = side == MergeGridSide.Merged ? CreateMergedPayloads(row) : GetSidePayloads(row, side);
        var recordExists = payloads is not null;
        var rowHeader = side switch
        {
            MergeGridSide.Base => FormatRowNumber(row.BaseRowNumber),
            MergeGridSide.Local => FormatRowNumber(row.LocalRowNumber),
            MergeGridSide.Remote => FormatRowNumber(row.RemoteRowNumber),
            MergeGridSide.Merged => recordExists ? FormatMergedRowNumber(row) : "-",
            _ => "-"
        };
        var cells = new MergeGridCell[ColumnHeaders.Count];
        for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++)
        {
            var payload = payloads?[columnIndex] ?? CellPayload.Blank;
            var conflictId = row.RowConflictId ?? row.CellConflictIds.GetValueOrDefault(columnIndex);
            var state = GetCellState(row, side, columnIndex, payloads, conflictId);
            cells[columnIndex] = new MergeGridCell(FormatCell(payload), state, conflictId);
        }
        return new MergeGridRow(rowHeader, row.RecordKey, cells);
    }

    private CellPayload[]? CreateMergedPayloads(ComparisonRowPlan row)
    {
        if (row.IsStructure)
        {
            var metadataResult = row.LocalCells?.ToArray() ?? new CellPayload[ColumnHeaders.Count];
            for (var columnIndex = 0; columnIndex < metadataResult.Length; columnIndex++)
            {
                if (!_mergedColumnIncluded(columnIndex))
                {
                    metadataResult[columnIndex] = CellPayload.Blank;
                    continue;
                }
                if (row.CellConflictIds.TryGetValue(columnIndex, out var conflictId))
                {
                    var choice = _conflicts[conflictId].SelectedChoice ?? MergeChoice.Local;
                    metadataResult[columnIndex] = GetChoicePayloads(row, choice)?[columnIndex] ?? CellPayload.Blank;
                }
                else
                {
                    metadataResult[columnIndex] = MergeNonConflictingCell(row, columnIndex);
                }
            }
            return metadataResult;
        }

        if (row.RowConflictId is { } rowConflictId)
        {
            var choice = _conflicts[rowConflictId].SelectedChoice ?? MergeChoice.Local;
            return GetChoicePayloads(row, choice);
        }

        var recordExists = DetermineMergedRecordPresence(row);
        if (!recordExists)
            return null;

        var result = new CellPayload[ColumnHeaders.Count];
        for (var columnIndex = 0; columnIndex < result.Length; columnIndex++)
        {
            if (!_mergedColumnIncluded(columnIndex))
            {
                result[columnIndex] = CellPayload.Blank;
                continue;
            }
            if (row.CellConflictIds.TryGetValue(columnIndex, out var conflictId))
            {
                var choice = _conflicts[conflictId].SelectedChoice ?? MergeChoice.Local;
                result[columnIndex] = GetChoicePayloads(row, choice)?[columnIndex] ?? CellPayload.Blank;
                continue;
            }

            result[columnIndex] = MergeNonConflictingCell(row, columnIndex);
        }
        return result;
    }

    private bool DetermineMergedRecordPresence(ComparisonRowPlan row)
    {
        if (row.BaseCells is null)
            return row.LocalCells is not null || row.RemoteCells is not null;
        if (row.LocalCells is null && row.RemoteCells is null)
            return false;
        if (row.LocalCells is null || row.RemoteCells is null)
            return false;
        return true;
    }

    private CellPayload MergeNonConflictingCell(ComparisonRowPlan row, int columnIndex)
    {
        var baseCell = row.BaseCells?[columnIndex] ?? CellPayload.Blank;
        var localCell = row.LocalCells?[columnIndex] ?? CellPayload.Blank;
        var remoteCell = row.RemoteCells?[columnIndex] ?? CellPayload.Blank;

        if (row.BaseCells is null)
            return row.LocalCells is not null ? localCell : remoteCell;
        if (row.LocalCells is null)
            return CellPayload.Blank;
        if (row.RemoteCells is null)
            return localCell;
        if (_ignoredColumns.Contains(columnIndex))
            return localCell;

        var structuralResult = _structuralCellMerge(
            columnIndex,
            row.RecordKey,
            baseCell,
            localCell,
            remoteCell);
        if (structuralResult is not null)
            return structuralResult;

        var decision = CellThreeWayMerger.Merge(baseCell, localCell, remoteCell, row.RecordKey, ColumnName(columnIndex));
        return decision.Result ?? localCell;
    }

    private MergeGridCellState GetCellState(
        ComparisonRowPlan row,
        MergeGridSide side,
        int columnIndex,
        CellPayload[]? payloads,
        string? conflictId)
    {
        if (row.IsStructure)
        {
            if (conflictId is not null)
                return MergeGridCellState.Metadata;
            var basePayload = row.BaseCells?[columnIndex] ?? CellPayload.Blank;
            var currentPayload = payloads?[columnIndex] ?? CellPayload.Blank;
            if (basePayload.Kind == CellValueKind.Blank && currentPayload.Kind != CellValueKind.Blank)
                return MergeGridCellState.Added;
            return currentPayload.ContentEquals(basePayload)
                ? MergeGridCellState.Normal
                : MergeGridCellState.Modified;
        }

        if (payloads is null)
        {
            return side switch
            {
                MergeGridSide.Local when row.BaseCells is not null => MergeGridCellState.Deleted,
                MergeGridSide.Remote when row.BaseCells is not null => MergeGridCellState.Deleted,
                MergeGridSide.Merged when row.BaseCells is not null => MergeGridCellState.Deleted,
                _ => MergeGridCellState.Normal
            };
        }

        if (conflictId is not null &&
            (side != MergeGridSide.Merged || !_conflicts[conflictId].IsResolved))
            return MergeGridCellState.Conflict;

        if (row.BaseCells is null)
            return MergeGridCellState.Added;

        var payload = payloads[columnIndex];
        return payload.ContentEquals(row.BaseCells[columnIndex])
            ? MergeGridCellState.Normal
            : MergeGridCellState.Modified;
    }

    private static CellPayload[]? GetSidePayloads(ComparisonRowPlan row, MergeGridSide side) => side switch
    {
        MergeGridSide.Base => row.BaseCells,
        MergeGridSide.Local => row.LocalCells,
        MergeGridSide.Remote => row.RemoteCells,
        _ => throw new ArgumentOutOfRangeException(nameof(side))
    };

    private static CellPayload[]? GetChoicePayloads(ComparisonRowPlan row, MergeChoice choice) => choice switch
    {
        MergeChoice.Base => row.BaseCells,
        MergeChoice.Local => row.LocalCells,
        MergeChoice.Remote => row.RemoteCells,
        _ => throw new ArgumentOutOfRangeException(nameof(choice))
    };

    private static bool RecordsEqual(CellPayload[]? left, CellPayload[]? right)
    {
        if (left is null || right is null)
            return left is null && right is null;
        return left.Length == right.Length && left.Zip(right).All(pair => pair.First.ContentEquals(pair.Second));
    }

    private static string FormatMergedRowNumber(ComparisonRowPlan row) =>
        row.LocalRowNumber is { } local ? local.ToString() : "+";

    private static string FormatRowNumber(int? rowNumber) => rowNumber?.ToString() ?? "-";

    internal static string FormatCell(CellPayload payload) => payload.Kind switch
    {
        CellValueKind.Blank => string.Empty,
        CellValueKind.Formula => $"={payload.FormulaText}",
        _ => payload.RawValue ?? string.Empty
    };

    internal static string ColumnName(int columnIndex)
    {
        var value = columnIndex + 1;
        var name = string.Empty;
        while (value > 0)
        {
            value--;
            name = (char)('A' + value % 26) + name;
            value /= 26;
        }
        return name;
    }
}

internal sealed class LazyReadOnlyList<T> : IReadOnlyList<T>
    where T : class
{
    private readonly T?[] _items;
    private readonly Func<int, T> _factory;

    public LazyReadOnlyList(int count, Func<int, T> factory)
    {
        _items = new T?[count];
        _factory = factory;
    }

    public int Count => _items.Length;

    public T this[int index]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            if (index >= _items.Length)
                throw new ArgumentOutOfRangeException(nameof(index));
            return _items[index] ??= _factory(index);
        }
    }

    public IEnumerator<T> GetEnumerator()
    {
        for (var index = 0; index < Count; index++)
            yield return this[index];
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}

internal sealed record ComparisonRowPlan(
    bool IsStructure,
    string RecordKey,
    int? BaseRowNumber,
    int? LocalRowNumber,
    int? RemoteRowNumber,
    CellPayload[]? BaseCells,
    CellPayload[]? LocalCells,
    CellPayload[]? RemoteCells,
    string? RowConflictId,
    IReadOnlyDictionary<int, string> CellConflictIds);
