namespace LubanExcelMerge.Core;

public sealed class LubanRecord
{
    private readonly IReadOnlyDictionary<string, CellPayload> _fields;

    public LubanRecord(string key, IEnumerable<KeyValuePair<string, CellPayload>> fields)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        Key = key;
        _fields = new Dictionary<string, CellPayload>(fields, StringComparer.Ordinal);
    }

    public string Key { get; }
    public IReadOnlyDictionary<string, CellPayload> Fields => _fields;

    public bool ContentEquals(LubanRecord? other)
    {
        if (other is null || _fields.Count != other._fields.Count)
            return false;

        return _fields.All(pair =>
            other._fields.TryGetValue(pair.Key, out var value) && pair.Value.ContentEquals(value));
    }
}
