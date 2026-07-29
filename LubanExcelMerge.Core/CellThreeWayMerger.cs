namespace LubanExcelMerge.Core;

public static class CellThreeWayMerger
{
    public static MergeDecision<CellPayload> Merge(
        CellPayload @base,
        CellPayload local,
        CellPayload remote,
        string recordKey,
        string fieldName)
    {
        var localEqualsBase = local.ContentEquals(@base);
        var remoteEqualsBase = remote.ContentEquals(@base);

        if (localEqualsBase && remoteEqualsBase)
            return new(MergeDecisionKind.Unchanged, local);
        if (localEqualsBase)
            return new(MergeDecisionKind.TakeRemote, remote);
        if (remoteEqualsBase)
            return new(MergeDecisionKind.TakeLocal, local);
        if (local.ContentEquals(remote))
            return new(MergeDecisionKind.BothChangedIdentically, local);

        return new(
            MergeDecisionKind.Conflict,
            null,
            new MergeConflict(
                MergeConflictKind.CellChangedDifferently,
                recordKey,
                fieldName,
                $"记录 {recordKey} 的字段 {fieldName} 在 LOCAL 和 REMOTE 中被修改为不同内容。"));
    }
}
