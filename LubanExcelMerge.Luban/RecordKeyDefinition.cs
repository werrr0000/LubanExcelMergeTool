namespace LubanExcelMerge.Luban;

public sealed record RecordKeyDefinition
{
    public RecordKeyDefinition(IEnumerable<string> fieldNames)
    {
        FieldNames = fieldNames.ToArray();
        if (FieldNames.Count == 0 || FieldNames.Any(string.IsNullOrEmpty))
            throw new ArgumentException("主键必须至少包含一个非空字段。", nameof(fieldNames));
    }

    public IReadOnlyList<string> FieldNames { get; }
    public string DisplayName => string.Join("+", FieldNames);

    public static IReadOnlyList<RecordKeyDefinition> ParseDeclaration(string declaration)
    {
        if (string.IsNullOrWhiteSpace(declaration))
            return Array.Empty<RecordKeyDefinition>();

        var normalized = new string(declaration.Where(character => !char.IsWhiteSpace(character)).ToArray());
        return normalized
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(candidate => new RecordKeyDefinition(candidate.Split('+')))
            .ToArray();
    }
}

public sealed record LubanRecordKey(IReadOnlyList<string> Components)
{
    public string StableValue => string.Concat(Components.Select(component => $"{component.Length}:{component}"));
    public string DisplayValue => string.Join(" | ", Components);
}
