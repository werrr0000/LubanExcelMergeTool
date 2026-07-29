namespace LubanExcelMerge.Luban;

public sealed record LubanRecordSource(
    string WorkbookPath,
    string SheetName,
    int RowNumber,
    IReadOnlyDictionary<string, string?> Values);

public enum KeyValidationIssueKind
{
    MissingField,
    EmptyKey,
    DuplicateKey
}

public sealed record KeyValidationIssue(
    KeyValidationIssueKind Kind,
    string WorkbookPath,
    string SheetName,
    int RowNumber,
    string KeyValue,
    string Message);

public sealed record RecordKeyValidationResult(
    RecordKeyDefinition KeyDefinition,
    IReadOnlyList<KeyValidationIssue> Issues)
{
    public bool IsValid => Issues.Count == 0;
}

public static class RecordKeyValidator
{
    public static RecordKeyValidationResult Validate(
        LubanSchema schema,
        RecordKeyDefinition keyDefinition,
        IReadOnlyList<LubanRecordSource> records)
    {
        var issues = new List<KeyValidationIssue>();
        foreach (var fieldName in keyDefinition.FieldNames)
        {
            if (schema.FindField(fieldName) is null)
            {
                issues.Add(new KeyValidationIssue(
                    KeyValidationIssueKind.MissingField,
                    records.FirstOrDefault()?.WorkbookPath ?? string.Empty,
                    records.FirstOrDefault()?.SheetName ?? string.Empty,
                    0,
                    string.Empty,
                    $"主键字段 {fieldName} 不存在，字段名区分大小写。"));
            }
        }

        if (issues.Count > 0)
            return new RecordKeyValidationResult(keyDefinition, issues);

        var keyedRecords = new List<(LubanRecordSource Record, LubanRecordKey Key)>();
        foreach (var record in records)
        {
            var components = keyDefinition.FieldNames.Select(fieldName =>
            {
                record.Values.TryGetValue(fieldName, out var value);
                return LubanKeyValueNormalizer.Normalize(schema.FindField(fieldName)!, value);
            }).ToArray();
            if (components.Any(string.IsNullOrEmpty))
            {
                issues.Add(new KeyValidationIssue(
                    KeyValidationIssueKind.EmptyKey,
                    record.WorkbookPath,
                    record.SheetName,
                    record.RowNumber,
                    string.Empty,
                    $"第 {record.RowNumber} 行的主键 {keyDefinition.DisplayName} 为空或不完整。"));
                continue;
            }

            keyedRecords.Add((record, new LubanRecordKey(components!)));
        }

        var duplicates = keyedRecords
            .GroupBy(item => item.Key.StableValue, StringComparer.Ordinal)
            .Where(group => group.Count() > 1);
        foreach (var duplicate in duplicates)
        {
            foreach (var item in duplicate)
            {
                issues.Add(new KeyValidationIssue(
                    KeyValidationIssueKind.DuplicateKey,
                    item.Record.WorkbookPath,
                    item.Record.SheetName,
                    item.Record.RowNumber,
                    item.Key.DisplayValue,
                    $"第 {item.Record.RowNumber} 行的主键 {item.Key.DisplayValue} 重复。"));
            }
        }

        return new RecordKeyValidationResult(keyDefinition, issues);
    }
}
