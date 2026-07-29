namespace LubanExcelMerge.Luban;

public sealed record LubanRawRow(int RowNumber, IReadOnlyList<string?> Cells);

public sealed record LubanField(int ColumnIndex, string Name, string TypeName);

public sealed record LubanSchema(
    IReadOnlyList<LubanField> Fields,
    IReadOnlyList<LubanRawRow> MetadataRows,
    int PrimaryVariableRowNumber,
    int? TypeRowNumber,
    int DataStartRowNumber,
    bool IsRestricted,
    IReadOnlyList<string> Restrictions)
{
    private readonly IReadOnlyDictionary<string, LubanField> _fieldsByName = Fields
        .GroupBy(field => field.Name, StringComparer.Ordinal)
        .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

    public LubanField? FindField(string name) => _fieldsByName.GetValueOrDefault(name);
}

public static class LubanSchemaParser
{
    public static LubanSchema Parse(IReadOnlyList<LubanRawRow> rows)
    {
        ArgumentNullException.ThrowIfNull(rows);
        if (rows.Count == 0)
            throw new FormatException("工作表为空，无法识别 Luban 表头。");

        var metadataRows = rows.TakeWhile(IsMetadataRow).ToArray();
        if (metadataRows.Length == 0)
            throw new FormatException("工作表顶部没有连续的 Luban 元数据区。");

        var variableRows = metadataRows.Where(row => GetMarker(row) == "##var").ToArray();
        if (variableRows.Length == 0)
            throw new FormatException("Luban 元数据区缺少 ##var 字段行。");

        var primaryVariableRow = variableRows[0];
        var typeRow = metadataRows.FirstOrDefault(row => GetMarker(row) == "##type");
        var fields = new List<LubanField>();
        for (var columnIndex = 0; columnIndex < primaryVariableRow.Cells.Count; columnIndex++)
        {
            var name = primaryVariableRow.Cells[columnIndex] ?? string.Empty;
            if (string.IsNullOrEmpty(name) || name.StartsWith("##", StringComparison.Ordinal))
                continue;

            var typeName = typeRow is not null && columnIndex < typeRow.Cells.Count
                ? typeRow.Cells[columnIndex] ?? string.Empty
                : string.Empty;
            fields.Add(new LubanField(columnIndex, name, typeName));
        }

        var restrictions = new List<string>();
        var duplicates = fields.GroupBy(field => field.Name, StringComparer.Ordinal).Where(group => group.Count() > 1);
        foreach (var duplicate in duplicates)
            restrictions.Add($"字段名 {duplicate.Key} 在主 ##var 行中重复。");

        if (fields.Count == 0)
            restrictions.Add("主 ##var 行没有可导出的字段。");

        return new LubanSchema(
            fields,
            metadataRows,
            primaryVariableRow.RowNumber,
            typeRow?.RowNumber,
            metadataRows[^1].RowNumber + 1,
            restrictions.Count > 0,
            restrictions);
    }

    private static bool IsMetadataRow(LubanRawRow row) => GetMarker(row).StartsWith("##", StringComparison.Ordinal);

    private static string GetMarker(LubanRawRow row) =>
        row.Cells.FirstOrDefault(cell => !string.IsNullOrEmpty(cell)) ?? string.Empty;
}
