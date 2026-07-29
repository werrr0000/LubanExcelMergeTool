using LubanExcelMerge.Git;

namespace LubanExcelMerge.Cli;

internal static class GitConfigCommand
{
    internal static int Run(IReadOnlyList<string> args, TextWriter output)
    {
        var values = Parse(args);
        var configuration = new GitMergeToolConfiguration(
            Required(values, "--gui"),
            Required(values, "--repo-root"),
            values.GetValueOrDefault("--tool-name", "LubanExcelMerge"));
        try
        {
            configuration.Validate();
        }
        catch (Exception exception) when (exception is ArgumentException or IOException)
        {
            throw new MergeInputException(exception.Message, exception);
        }

        output.WriteLine("以下内容仅供确认，本命令不会修改 Git 配置。");
        output.WriteLine();
        output.WriteLine(configuration.BuildIniSnippet());
        output.WriteLine();
        output.WriteLine("确认后可在 PowerShell 中执行以下仓库级命令：");
        foreach (var command in configuration.BuildLocalConfigurationCommands())
            output.WriteLine(command);
        return ExitCodes.Success;
    }

    private static Dictionary<string, string> Parse(IReadOnlyList<string> args)
    {
        var supported = new HashSet<string>(StringComparer.Ordinal)
        {
            "--gui", "--repo-root", "--tool-name"
        };
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 1; index < args.Count; index++)
        {
            var option = args[index];
            if (!supported.Contains(option))
                throw new MergeInputException($"git-config 的未知参数：{option}。");
            if (values.ContainsKey(option))
                throw new MergeInputException($"参数 {option} 重复。");
            if (++index >= args.Count || args[index].StartsWith("--", StringComparison.Ordinal))
                throw new MergeInputException($"参数 {option} 缺少值。");
            values[option] = args[index];
        }
        return values;
    }

    private static string Required(IReadOnlyDictionary<string, string> values, string option) =>
        values.TryGetValue(option, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new MergeInputException($"git-config 缺少必需参数 {option}。");
}
