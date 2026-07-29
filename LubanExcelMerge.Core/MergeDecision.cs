namespace LubanExcelMerge.Core;

public enum MergeDecisionKind
{
    Unchanged,
    TakeLocal,
    TakeRemote,
    BothChangedIdentically,
    AddedLocal,
    AddedRemote,
    AddedIdentically,
    Deleted,
    Conflict
}

public enum MergeConflictKind
{
    CellChangedDifferently,
    AddAdd,
    DeleteModify,
    MetadataChanged
}

public sealed record MergeConflict(
    MergeConflictKind Kind,
    string RecordKey,
    string? FieldName,
    string Message);

public sealed record MergeDecision<T>(
    MergeDecisionKind Kind,
    T? Result,
    MergeConflict? Conflict = null)
{
    public bool IsConflict => Conflict is not null;
}
