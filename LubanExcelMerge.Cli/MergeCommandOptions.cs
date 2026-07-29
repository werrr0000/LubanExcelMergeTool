namespace LubanExcelMerge.Cli;

public sealed record MergeCommandOptions(
    string BasePath,
    string LocalPath,
    string RemotePath,
    string OutputPath,
    string RepositoryRoot,
    string? DataRoot,
    string? TablesPath,
    string? ConfigPath,
    string Language,
    bool Headless,
    bool Validate,
    string? RecalculateWithExcel,
    string? LogPath)
{
    public string? LoadedConfigPath { get; init; }
    public string? ProjectValidationCommand { get; init; }
    public bool ProjectValidationEnabled { get; init; }
    public bool ValidateLogicalTableUniqueness { get; init; }
    public IReadOnlyDictionary<string, string[]>? KeyOverrides { get; init; }
    public IReadOnlyDictionary<string, string[]>? IgnoredFields { get; init; }
    public IReadOnlyList<string>? InactivePaths { get; init; }
    public bool FullExportValidationEnabled { get; init; }
    public string? FullExportValidationCommand { get; init; }
}

public static class CommandLineParser
{
    private static readonly HashSet<string> ValueOptions = new(StringComparer.Ordinal)
    {
        "--base", "--local", "--remote", "--output", "--repo-root", "--data-root", "--tables",
        "--config", "--language", "--recalculate-with-excel", "--log"
    };

    private static readonly HashSet<string> SwitchOptions = new(StringComparer.Ordinal)
    {
        "--headless", "--validate", "--validate-full"
    };

    public static MergeCommandOptions Parse(IReadOnlyList<string> args)
    {
        var argumentStart = GetMergeArgumentStart(args);

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var switches = new HashSet<string>(StringComparer.Ordinal);
        for (var index = argumentStart; index < args.Count; index++)
        {
            var argument = args[index];
            if (SwitchOptions.Contains(argument))
            {
                if (!switches.Add(argument))
                    throw new MergeInputException($"参数 {argument} 重复。");
                continue;
            }

            if (!ValueOptions.Contains(argument))
                throw new MergeInputException($"未知参数：{argument}。");
            if (values.ContainsKey(argument))
                throw new MergeInputException($"参数 {argument} 重复。");
            if (++index >= args.Count || args[index].StartsWith("--", StringComparison.Ordinal))
                throw new MergeInputException($"参数 {argument} 缺少值。");
            values[argument] = args[index];
        }

        var language = values.GetValueOrDefault("--language", "zh-CN");
        if (!string.Equals(language, "zh-CN", StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("当前版本只支持 --language zh-CN。");
        var recalculation = values.GetValueOrDefault("--recalculate-with-excel");
        if (recalculation is not null and not ("auto" or "always" or "never"))
            throw new MergeInputException("--recalculate-with-excel 必须是 auto、always 或 never。");

        return new MergeCommandOptions(
            Required(values, "--base"),
            Required(values, "--local"),
            Required(values, "--remote"),
            Required(values, "--output"),
            Required(values, "--repo-root"),
            values.GetValueOrDefault("--data-root"),
            values.GetValueOrDefault("--tables"),
            values.GetValueOrDefault("--config"),
            language,
            switches.Contains("--headless"),
            switches.Contains("--validate"),
            recalculation,
            values.GetValueOrDefault("--log"))
        {
            FullExportValidationEnabled = switches.Contains("--validate-full")
        };
    }

    public static string Usage =>
        "LubanExcelMerge merge --base <BASE.xlsx> --local <LOCAL.xlsx> --remote <REMOTE.xlsx> " +
        "--output <MERGED.xlsx> --repo-root <仓库根目录> [--data-root <Datas>] [--tables <__tables__.csv>] " +
        "[--config <luban-excel-merge.json>] [--recalculate-with-excel <auto|always|never>] " +
        "[--validate] [--validate-full] [--log <日志路径>] [--headless]" +
        Environment.NewLine +
        "LubanExcelMerge git-config --gui <LubanExcelMerge.Gui.exe> --repo-root <仓库根目录> [--tool-name <名称>]" +
        Environment.NewLine +
        "LubanExcelMerge diagnostic-package --log <日志.jsonl> --output <诊断包.zip>" +
        Environment.NewLine +
        "Fork GUI Arguments 可省略 merge，直接从 --base 开始。";

    private static int GetMergeArgumentStart(IReadOnlyList<string> args)
    {
        if (args.Count == 0)
            throw new MergeInputException("缺少 merge 子命令或合并参数。");
        if (string.Equals(args[0], "merge", StringComparison.Ordinal))
            return 1;
        if (ValueOptions.Contains(args[0]) || SwitchOptions.Contains(args[0]))
            return 0;

        throw new MergeInputException($"未知子命令：{args[0]}。合并请使用 merge，Fork GUI Arguments 也可直接从 --base 开始。");
    }

    private static string Required(IReadOnlyDictionary<string, string> values, string name) =>
        values.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new MergeInputException($"缺少必需参数 {name}。");
}
