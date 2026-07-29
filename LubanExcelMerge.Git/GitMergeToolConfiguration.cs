namespace LubanExcelMerge.Git;

public sealed record GitMergeToolConfiguration(
    string GuiExecutablePath,
    string RepositoryRoot,
    string ToolName = "LubanExcelMerge")
{
    public string Command =>
        $"{ShellQuote(NormalizePath(GuiExecutablePath))} merge " +
        "--base \"$BASE\" --local \"$LOCAL\" --remote \"$REMOTE\" --output \"$MERGED\" " +
        $"--repo-root {ShellQuote(NormalizePath(RepositoryRoot))} --recalculate-with-excel auto";

    public string BuildIniSnippet()
    {
        var escapedCommand = Command.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal);
        return $"""
            [merge]
                tool = {ToolName}

            [mergetool "{ToolName}"]
                cmd = "{escapedCommand}"
                trustExitCode = true
            """;
    }

    public IReadOnlyList<string> BuildLocalConfigurationCommands() =>
    [
        $"git -C {PowerShellQuote(Path.GetFullPath(RepositoryRoot))} config --local merge.tool {PowerShellQuote(ToolName)}",
        $"git -C {PowerShellQuote(Path.GetFullPath(RepositoryRoot))} config --local mergetool.{ToolName}.cmd {PowerShellQuote(Command)}",
        $"git -C {PowerShellQuote(Path.GetFullPath(RepositoryRoot))} config --local mergetool.{ToolName}.trustExitCode true"
    ];

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(ToolName) || ToolName.Any(char.IsWhiteSpace))
            throw new ArgumentException("Git merge tool 名称不能为空或包含空白。", nameof(ToolName));
        if (!File.Exists(GuiExecutablePath))
            throw new FileNotFoundException("LubanExcelMerge GUI 可执行文件不存在。", GuiExecutablePath);
        if (!Directory.Exists(RepositoryRoot) ||
            !Directory.Exists(Path.Combine(RepositoryRoot, ".git")) &&
            !File.Exists(Path.Combine(RepositoryRoot, ".git")))
            throw new DirectoryNotFoundException($"{RepositoryRoot} 不是 Git 仓库根目录。");
    }

    private static string NormalizePath(string path) => Path.GetFullPath(path).Replace('\\', '/');

    private static string ShellQuote(string value) => "'" + value.Replace("'", "'\"'\"'", StringComparison.Ordinal) + "'";

    private static string PowerShellQuote(string value) =>
        "'" + value.Replace("'", "''", StringComparison.Ordinal) + "'";
}
