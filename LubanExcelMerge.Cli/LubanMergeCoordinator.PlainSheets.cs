using LubanExcelMerge.Core;
using LubanExcelMerge.Luban;
using LubanExcelMerge.OpenXml;
using System.Text.RegularExpressions;

namespace LubanExcelMerge.Cli;

public sealed partial class LubanMergeCoordinator
{
    // Recognize ordinary keyed tables conservatively. Descriptive pages without
    // an explicit Id header continue through fixed-address review, never key inference.
    private static bool TryParsePlainSchemas(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        out (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas)
    {
        schemas = default;
        var parsed = new[] { sheets.Base, sheets.Local, sheets.Remote }.Select(ParsePlainSchema).ToArray();
        if (parsed.Any(schema => schema is null)) return false;
        var first = parsed[0]!;
        if (parsed.Skip(1).Any(schema => schema!.PrimaryVariableRowNumber != first.PrimaryVariableRowNumber ||
                !schema.Fields.SequenceEqual(first.Fields)))
            throw new UnsafeWorkbookException($"普通工作表 {sheets.Local.Name} 的字段行或列布局不同，请先统一三方表头。");
        schemas = (parsed[0]!, parsed[1]!, parsed[2]!);
        return true;
    }

    private static LubanSchema? ParsePlainSchema(SheetSnapshot sheet)
    {
        if (HasSchemaMarker(sheet)) return null;
        foreach (var row in sheet.Rows.Where(row => row.RowNumber <= 8).OrderBy(row => row.RowNumber))
        {
            var cells = row.Cells.Where(HasCellContent).ToArray();
            if (cells.Length < 2 || !cells.Any(cell => cell.Payload.RawValue is "Id" or "id") ||
                cells.Any(cell => cell.Payload.Kind != CellValueKind.String ||
                    !Regex.IsMatch(cell.Payload.RawValue ?? "", "^[A-Za-z_][A-Za-z0-9_]*$"))) continue;
            var fields = cells.Select(cell => new LubanField(cell.ColumnIndex, cell.Payload.RawValue!, "")).ToArray();
            if (fields.DistinctBy(field => field.Name).Count() != fields.Length) return null;
            var keys = fields.Where(field => field.Name is "Id" or "id").ToArray();
            if (keys.Length != 1) return null;
            var key = keys[0];
            var data = sheet.Rows.Where(dataRow => dataRow.RowNumber > row.RowNumber && dataRow.Cells.Any(HasCellContent)).ToArray();
            if (data.Length == 0 || data.Any(dataRow => string.IsNullOrWhiteSpace(dataRow.GetCell(key.ColumnIndex)?.Payload.RawValue)))
                return null;
            var columns = fields.Select(field => field.ColumnIndex).ToHashSet();
            if (data.Any(dataRow => dataRow.Cells.Any(cell => HasCellContent(cell) && !columns.Contains(cell.ColumnIndex))))
                return null;
            // Place Id first for stable key inference, retaining every physical column index.
            fields = fields.OrderBy(field => field == key ? 0 : 1).ThenBy(field => field.ColumnIndex).ToArray();
            var width = cells.Max(cell => cell.ColumnIndex) + 1;
            var prefix = Enumerable.Range(1, row.RowNumber).Select(number =>
            {
                var values = new string?[width];
                foreach (var cell in sheet.GetRow(number)?.Cells ?? Array.Empty<OpenXmlCellSnapshot>())
                    if (cell.ColumnIndex < width) values[cell.ColumnIndex] = cell.Payload.RawValue;
                return new LubanRawRow(number, values);
            }).ToArray();
            return new LubanSchema(fields, prefix, row.RowNumber, null, row.RowNumber + 1, false, Array.Empty<string>());
        }
        return null;
    }

    private static PreparedSheetMerge PrepareAuxiliarySheet(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        WorkbookSnapshot baseBook, WorkbookSnapshot localBook, WorkbookSnapshot remoteBook)
    {
        var name = sheets.Local.Name;
        var allCells = new[] { sheets.Base, sheets.Local, sheets.Remote }
            .SelectMany(sheet => sheet.Rows).SelectMany(row => row.Cells).Where(HasCellContent).ToArray();
        var width = allCells.Select(cell => cell.ColumnIndex).DefaultIfEmpty(0).Max() + 1;
        var rowNumbers = allCells.Select(cell => cell.RowNumber).Append(1).Distinct().Order().ToArray();
        var rows = new List<ComparisonRowPlan>();
        var conflicts = new List<ResolvableMergeConflict>();
        foreach (var number in rowNumbers)
        {
            CellPayload[] Values(SheetSnapshot sheet) => Enumerable.Range(0, width)
                .Select(column => sheet.GetRow(number)?.GetCell(column)?.Payload ?? CellPayload.Blank).ToArray();
            var b = Values(sheets.Base); var l = Values(sheets.Local); var r = Values(sheets.Remote);
            var ids = new Dictionary<int, string>();
            for (var column = 0; column < width; column++)
            {
                if (b[column].ContentEquals(l[column]) && b[column].ContentEquals(r[column])) continue;
                var address = CellReference.Create(number, column);
                IReadOnlyList<WorkbookEdit> Choice(SheetSnapshot source, CellPayload target) =>
                    l[column].ContentEquals(target) ? Array.Empty<WorkbookEdit>() :
                    new WorkbookEdit[] { new SetCellEdit(name, address, target, source.GetCell(address)?.StyleIndex)
                        { StyleSourceWorkbook = source.GetCell(address)?.Payload.SourceWorkbook } };
                var conflict = CreateResolution(conflicts.Count,
                    new MergeConflict(MergeConflictKind.MetadataChanged, $"辅助页 {name} 第{number}行", address,
                        $"辅助页 {name}!{address} 有变化，请按原单元格地址复核。工具不推断该页的行列移动。"),
                    number, FormatCell(b[column]), FormatCell(l[column]), FormatCell(r[column]),
                    Choice(sheets.Base, b[column]), Array.Empty<WorkbookEdit>(), Choice(sheets.Remote, r[column]));
                conflict.SetGridLocation(rows.Count, column);
                conflicts.Add(conflict); ids[column] = conflict.Id;
            }
            rows.Add(new ComparisonRowPlan(true, $"辅助页 {name} 第{number}行", number, number, number, b, l, r, null, ids));
        }
        var notices = new List<string>();
        if (!AuxiliaryPartsEqual(baseBook, sheets.Base, localBook, sheets.Local) ||
            !AuxiliaryPartsEqual(baseBook, sheets.Base, remoteBook, sheets.Remote))
            notices.Add("辅助页保留 LOCAL 的行列布局、样式、批注和图形；所选内容单元格导入来源样式，请确认该布局可接受");
        if (conflicts.Count > 0)
            notices.Add("辅助页按固定单元格地址复核，不自动推断记录身份或行列移动");
        return new PreparedSheetMerge(name, "辅助页：单元格地址", Array.Empty<WorkbookEdit>(), conflicts,
            new MergeComparison(Enumerable.Range(0, width).Select(MergeComparison.ColumnName).ToArray(), rows, conflicts, new HashSet<int>()),
            0, 0, 0, conflicts.Count, notices) { UsesFixedCellReview = true };
    }
}
