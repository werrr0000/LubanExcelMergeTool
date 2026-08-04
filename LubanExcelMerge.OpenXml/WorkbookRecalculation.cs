namespace LubanExcelMerge.OpenXml;

public enum WorkbookRecalculationMode
{
    Never,
    Auto,
    Always
}

public enum WorkbookRecalculationStatus
{
    NotNeeded,
    SourceCachePreservedUnverified,
    DeferredAfterRecalculationFailure,
    Completed
}

public sealed record WorkbookSaveOptions(
    WorkbookRecalculationMode RecalculationMode = WorkbookRecalculationMode.Never,
    bool FormulaMayBeAffected = false,
    TimeSpan? RecalculationTimeout = null)
{
    public TimeSpan EffectiveTimeout => RecalculationTimeout ??
        (RecalculationMode == WorkbookRecalculationMode.Auto ? TimeSpan.FromSeconds(30) : TimeSpan.FromMinutes(2));
}

public interface IWorkbookRecalculator
{
    string ProviderName { get; }
    bool IsAvailable { get; }
    void Recalculate(string workbookPath, TimeSpan timeout);
}

public sealed class WorkbookRecalculationUnavailableException : InvalidOperationException
{
    public WorkbookRecalculationUnavailableException(string message) : base(message) { }
}
