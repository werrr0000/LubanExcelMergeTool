using LubanExcelMerge.Core;

namespace LubanExcelMerge.OpenXml;

public abstract record WorkbookEdit(string SheetName);

public sealed record SetCellEdit(
    string SheetName,
    string Address,
    CellPayload Payload,
    string? StyleIndex = null) : WorkbookEdit(SheetName);

public sealed record DeleteRowEdit(string SheetName, int RowNumber) : WorkbookEdit(SheetName);

public sealed record CleanupEmptyFormattingEdit(string SheetName) : WorkbookEdit(SheetName);

public sealed record CellWrite(int ColumnIndex, CellPayload Payload, string? StyleIndex = null);

public sealed record AppendRowEdit(
    string SheetName,
    IReadOnlyList<CellWrite> Cells,
    int? SourceRowNumber = null) : WorkbookEdit(SheetName);

public sealed record RowWrite(
    IReadOnlyList<CellWrite> Cells);

public sealed record ReplaceMetadataRowsEdit(
    string SheetName,
    int StartRowNumber,
    int ExistingRowCount,
    IReadOnlyList<RowWrite> Rows) : WorkbookEdit(SheetName);
