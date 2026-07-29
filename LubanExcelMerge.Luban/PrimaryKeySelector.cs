namespace LubanExcelMerge.Luban;

public sealed record PrimaryKeySelection(
    RecordKeyDefinition? Selected,
    IReadOnlyList<RecordKeyValidationResult> Attempts)
{
    public bool IsValid => Selected is not null;
}

public static class PrimaryKeySelector
{
    public static PrimaryKeySelection Select(
        LubanSchema schema,
        IReadOnlyList<RecordKeyDefinition> declaredIndexes,
        params IReadOnlyList<LubanRecordSource>[] datasets)
    {
        var candidates = declaredIndexes.Count > 0
            ? declaredIndexes
            : schema.Fields.Take(1).Select(field => new RecordKeyDefinition(new[] { field.Name })).ToArray();
        var attempts = new List<RecordKeyValidationResult>();

        foreach (var candidate in candidates)
        {
            var candidateResults = datasets
                .Select(dataset => RecordKeyValidator.Validate(schema, candidate, dataset))
                .ToArray();
            attempts.AddRange(candidateResults);
            if (candidateResults.All(result => result.IsValid))
                return new PrimaryKeySelection(candidate, attempts);
        }

        return new PrimaryKeySelection(null, attempts);
    }
}
