using LubanExcelMerge.Core;

namespace LubanExcelMerge.OpenXml;

public sealed record OpenXmlCellSnapshot(
    string Address,
    int RowNumber,
    int ColumnIndex,
    string? StyleIndex,
    CellPayload Payload);

public sealed class OpenXmlRowSnapshot
{
    private readonly OpenXmlCellSnapshot[] _cells;

    public OpenXmlRowSnapshot(int rowNumber, IReadOnlyList<OpenXmlCellSnapshot> cells)
    {
        RowNumber = rowNumber;
        _cells = cells.OrderBy(cell => cell.ColumnIndex).ToArray();
    }

    public int RowNumber { get; }
    public IReadOnlyList<OpenXmlCellSnapshot> Cells => _cells;

    public OpenXmlCellSnapshot? GetCell(int columnIndex)
    {
        var low = 0;
        var high = _cells.Length - 1;
        while (low <= high)
        {
            var middle = low + (high - low) / 2;
            var cell = _cells[middle];
            if (cell.ColumnIndex == columnIndex)
                return cell;
            if (cell.ColumnIndex < columnIndex)
                low = middle + 1;
            else
                high = middle - 1;
        }

        return null;
    }
}

public sealed class SheetSnapshot
{
    private readonly IReadOnlyDictionary<string, OpenXmlCellSnapshot> _cells;
    private readonly IReadOnlyDictionary<int, OpenXmlRowSnapshot> _rowsByNumber;
    private readonly int _formulaCount;

    public SheetSnapshot(string name, string partPath, IReadOnlyList<OpenXmlRowSnapshot> rows)
    {
        Name = name;
        PartPath = partPath;
        Rows = rows;
        _cells = rows.SelectMany(row => row.Cells).ToDictionary(cell => cell.Address, StringComparer.OrdinalIgnoreCase);
        _rowsByNumber = rows.ToDictionary(row => row.RowNumber);
        _formulaCount = _cells.Values.Count(cell => cell.Payload.Kind == CellValueKind.Formula);
    }

    public string Name { get; }
    public string PartPath { get; }
    public IReadOnlyList<OpenXmlRowSnapshot> Rows { get; }
    public int FormulaCount => _formulaCount;

    public OpenXmlCellSnapshot? GetCell(string address) =>
        _cells.TryGetValue(address, out var cell) ? cell : null;

    public OpenXmlRowSnapshot? GetRow(int rowNumber) =>
        _rowsByNumber.GetValueOrDefault(rowNumber);
}

public sealed record WorkbookSnapshot(
    string SourcePath,
    IReadOnlyList<SheetSnapshot> Sheets,
    IReadOnlyList<string> PackagePartNames)
{
    public SheetSnapshot GetSheet(string name) => Sheets.FirstOrDefault(sheet =>
        string.Equals(sheet.Name, name, StringComparison.Ordinal))
        ?? throw new KeyNotFoundException($"工作簿中不存在工作表 {name}。");
}

public sealed record OpenXmlReadLimits(
    int MaxEntries = 10_000,
    long MaxTotalUncompressedBytes = 512L * 1024 * 1024,
    double MaxCompressionRatio = 1_000);
