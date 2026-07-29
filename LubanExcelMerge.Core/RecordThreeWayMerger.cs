namespace LubanExcelMerge.Core;

public sealed record RecordMergeResult(
    MergeDecisionKind Kind,
    LubanRecord? Record,
    IReadOnlyList<MergeConflict> Conflicts)
{
    public bool IsConflict => Conflicts.Count > 0;
}

public static class RecordThreeWayMerger
{
    public static RecordMergeResult Merge(LubanRecord? @base, LubanRecord? local, LubanRecord? remote)
    {
        var key = @base?.Key ?? local?.Key ?? remote?.Key
            ?? throw new ArgumentException("At least one record version is required.");

        EnsureSameKey(key, @base, local, remote);

        if (@base is null)
            return MergeAdded(key, local, remote);

        if (local is null && remote is null)
            return new(MergeDecisionKind.Deleted, null, Array.Empty<MergeConflict>());

        if (local is null)
            return remote!.ContentEquals(@base)
                ? new(MergeDecisionKind.Deleted, null, Array.Empty<MergeConflict>())
                : DeleteModifyConflict(key);

        if (remote is null)
            return local.ContentEquals(@base)
                ? new(MergeDecisionKind.Deleted, null, Array.Empty<MergeConflict>())
                : DeleteModifyConflict(key);

        var fieldNames = @base.Fields.Keys
            .Concat(local.Fields.Keys)
            .Concat(remote.Fields.Keys)
            .Distinct(StringComparer.Ordinal);
        var mergedFields = new Dictionary<string, CellPayload>(StringComparer.Ordinal);
        var conflicts = new List<MergeConflict>();

        foreach (var fieldName in fieldNames)
        {
            var baseCell = GetCell(@base, fieldName);
            var localCell = GetCell(local, fieldName);
            var remoteCell = GetCell(remote, fieldName);
            var decision = CellThreeWayMerger.Merge(baseCell, localCell, remoteCell, key, fieldName);

            if (decision.Conflict is not null)
                conflicts.Add(decision.Conflict);
            else
                mergedFields[fieldName] = decision.Result!;
        }

        return conflicts.Count > 0
            ? new(MergeDecisionKind.Conflict, null, conflicts)
            : new(MergeDecisionKind.BothChangedIdentically, new LubanRecord(key, mergedFields), conflicts);
    }

    private static RecordMergeResult MergeAdded(string key, LubanRecord? local, LubanRecord? remote)
    {
        if (local is null && remote is null)
            throw new ArgumentException("An added record must exist in LOCAL or REMOTE.");
        if (local is null)
            return new(MergeDecisionKind.AddedRemote, remote, Array.Empty<MergeConflict>());
        if (remote is null)
            return new(MergeDecisionKind.AddedLocal, local, Array.Empty<MergeConflict>());
        if (local.ContentEquals(remote))
            return new(MergeDecisionKind.AddedIdentically, local, Array.Empty<MergeConflict>());

        return new(
            MergeDecisionKind.Conflict,
            null,
            new[] { new MergeConflict(MergeConflictKind.AddAdd, key, null, $"LOCAL 和 REMOTE 新增了同键但内容不同的记录 {key}。") });
    }

    private static RecordMergeResult DeleteModifyConflict(string key) => new(
        MergeDecisionKind.Conflict,
        null,
        new[] { new MergeConflict(MergeConflictKind.DeleteModify, key, null, $"记录 {key} 在一侧被删除、另一侧被修改。") });

    private static CellPayload GetCell(LubanRecord record, string fieldName) =>
        record.Fields.TryGetValue(fieldName, out var value) ? value : CellPayload.Blank;

    private static void EnsureSameKey(string key, params LubanRecord?[] records)
    {
        if (records.Any(record => record is not null && !string.Equals(record.Key, key, StringComparison.Ordinal)))
            throw new ArgumentException("All record versions must have the same key.");
    }
}
