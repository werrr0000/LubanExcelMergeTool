namespace LubanExcelMerge.OpenXml;

public sealed class WpsComWorkbookRecalculator : IWorkbookRecalculator
{
    private readonly ExcelComWorkbookRecalculator _inner = new(
        "WPS 表格",
        new[] { "ket.Application", "et.Application" },
        new[] { "et" });

    public string ProviderName => _inner.ProviderName;
    public bool IsAvailable => _inner.IsAvailable;
    public void Recalculate(string workbookPath, TimeSpan timeout) => _inner.Recalculate(workbookPath, timeout);
}
