namespace LubanExcelMerge.OpenXml;

public sealed class CompositeWorkbookRecalculator : IWorkbookRecalculator
{
    private readonly IReadOnlyList<IWorkbookRecalculator> _providers;

    public CompositeWorkbookRecalculator(params IWorkbookRecalculator[] providers)
    {
        if (providers.Length == 0)
            throw new ArgumentException("至少需要一个工作簿重算后端。", nameof(providers));
        _providers = providers;
    }

    public string ProviderName => SelectedProvider?.ProviderName ??
        string.Join(" / ", _providers.Select(provider => provider.ProviderName));
    public bool IsAvailable => SelectedProvider is not null;

    public void Recalculate(string workbookPath, TimeSpan timeout)
    {
        var provider = SelectedProvider
            ?? throw new WorkbookRecalculationUnavailableException(
                $"{ProviderName} 的自动化接口均不可用。");
        provider.Recalculate(workbookPath, timeout);
    }

    public static CompositeWorkbookRecalculator CreateDefault() =>
        new(new WpsComWorkbookRecalculator(), new ExcelComWorkbookRecalculator());

    private IWorkbookRecalculator? SelectedProvider => _providers.FirstOrDefault(provider => provider.IsAvailable);
}
