namespace LubanExcelMerge.Luban;

public sealed record LogicalTableDefinition(
    string FullName,
    string ValueType,
    bool ReadSchemaFromFile,
    IReadOnlyList<string> InputFiles,
    IReadOnlyList<RecordKeyDefinition> DeclaredIndexes,
    string Mode,
    string Group);

public sealed class LogicalTableCatalog
{
    private static readonly string[] RequiredColumns =
    {
        "full_name", "value_type", "read_schema_from_file", "input", "index", "mode", "group"
    };

    private readonly IReadOnlyList<LogicalTableDefinition> _tables;

    public LogicalTableCatalog(IReadOnlyList<LogicalTableDefinition> tables) => _tables = tables;

    public IReadOnlyList<LogicalTableDefinition> Tables => _tables;

    public static LogicalTableCatalog Parse(TextReader reader)
    {
        var rows = CsvReader.ReadAll(reader);
        var headerIndex = FindHeader(rows);
        var header = rows[headerIndex];
        var columns = header
            .Select((name, index) => (name, index))
            .Where(item => !string.IsNullOrEmpty(item.name))
            .ToDictionary(item => item.name, item => item.index, StringComparer.Ordinal);

        var tables = new List<LogicalTableDefinition>();
        for (var rowIndex = headerIndex + 1; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            var marker = Get(row, 0);
            if (marker.StartsWith("##", StringComparison.Ordinal))
                continue;

            var fullName = Get(row, columns["full_name"]);
            var input = Get(row, columns["input"]);
            if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(input))
                continue;

            tables.Add(new LogicalTableDefinition(
                fullName,
                Get(row, columns["value_type"]),
                ParseBoolean(Get(row, columns["read_schema_from_file"]), rowIndex + 1),
                input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
                RecordKeyDefinition.ParseDeclaration(Get(row, columns["index"])),
                Get(row, columns["mode"]),
                Get(row, columns["group"])));
        }

        return new LogicalTableCatalog(tables);
    }

    public IReadOnlyList<LogicalTableDefinition> MatchInput(string relativePath)
    {
        var normalizedPath = NormalizePath(relativePath);
        return _tables
            .Where(table => table.InputFiles.Any(input =>
                string.Equals(NormalizePath(input), normalizedPath, StringComparison.OrdinalIgnoreCase)))
            .ToArray();
    }

    private static int FindHeader(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            if (Get(row, 0) != "##var")
                continue;

            if (RequiredColumns.All(required => row.Contains(required, StringComparer.Ordinal)))
                return index;
        }

        throw new FormatException("__tables__.csv 中未找到包含必需字段的 ##var 表头。");
    }

    private static bool ParseBoolean(string value, int rowNumber)
    {
        if (bool.TryParse(value, out var result))
            return result;
        if (string.IsNullOrWhiteSpace(value))
            return false;

        throw new FormatException($"__tables__.csv 第 {rowNumber} 行的 read_schema_from_file 不是布尔值。");
    }

    private static string Get(IReadOnlyList<string> row, int index) => index < row.Count ? row[index] : string.Empty;

    private static string NormalizePath(string path) => path.Trim().Replace('\\', '/').TrimStart('/');
}
