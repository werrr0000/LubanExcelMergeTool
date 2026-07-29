using LubanExcelMerge.Core;

namespace LubanExcelMerge.OpenXml;

public abstract record WorkbookEdit(string SheetName);

public sealed record SetCellEdit(
    string SheetName,
    string Address,
    CellPayload Payload,
    string? StyleIndex = null) : WorkbookEdit(SheetName);

public sealed record DeleteRowEdit(string SheetName, int RowNumber) : WorkbookEdit(SheetName);

public sealed record CellWrite(int ColumnIndex, CellPayload Payload, string? StyleIndex = null);

public sealed record AppendRowEdit(
    string SheetName,
    IReadOnlyList<CellWrite> Cells,
    int? SourceRowNumber = null) : WorkbookEdit(SheetName);
