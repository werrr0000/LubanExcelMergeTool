using System.Diagnostics;

namespace LubanExcelMerge.Git;

public sealed record GitStageResult(
    string RepositoryRoot,
    string RelativePath,
    bool WasUnmerged);

public sealed class GitStagingException : Exception
{
    public GitStagingException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}

public sealed class GitMergedFileStager
{
    private static readonly TimeSpan GitTimeout = TimeSpan.FromSeconds(30);

    public GitStageResult Stage(string repositoryRoot, string mergedPath)
    {
        var requestedRoot = Path.GetFullPath(repositoryRoot);
        var absoluteMerged = Path.GetFullPath(mergedPath);
        if (!File.Exists(absoluteMerged))
            throw new GitStagingException($"MERGED 文件不存在，无法加入 Git staged：{absoluteMerged}。");
        if (!IsWithinRoot(requestedRoot, absoluteMerged))
            throw new GitStagingException("MERGED 文件不在指定项目根目录内，已拒绝自动暂存。");

        var mergedDirectory = Path.GetDirectoryName(absoluteMerged)
            ?? throw new GitStagingException("MERGED 文件路径缺少目录。");
        var topLevel = RunGit(mergedDirectory, "rev-parse", "--show-toplevel").Output.Trim();
        if (string.IsNullOrWhiteSpace(topLevel))
            throw new GitStagingException($"无法从 MERGED 路径确定 Git 仓库根目录：{absoluteMerged}。");
        var absoluteRoot = Path.GetFullPath(topLevel);
        if (!IsWithinRoot(requestedRoot, absoluteRoot) || !IsWithinRoot(absoluteRoot, absoluteMerged))
            throw new GitStagingException("MERGED 所属 Git 仓库超出指定项目根目录，已拒绝自动暂存。");
        var relativePath = Path.GetRelativePath(absoluteRoot, absoluteMerged);
        if (Path.IsPathRooted(relativePath) ||
            relativePath.Equals("..", StringComparison.Ordinal) ||
            relativePath.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) ||
            relativePath.StartsWith($"..{Path.AltDirectorySeparatorChar}", StringComparison.Ordinal))
        {
            throw new GitStagingException("MERGED 文件不在指定 Git 仓库内，已拒绝自动暂存。");
        }

        var gitPath = relativePath.Replace(Path.DirectorySeparatorChar, '/');
        var tracked = RunGit(absoluteRoot, allowFailure: true, "ls-files", "--error-unmatch", "--", gitPath);
        if (tracked.ExitCode != 0)
            throw new GitStagingException($"MERGED 不是 Git 已跟踪文件，已拒绝自动暂存：{gitPath}。");

        var unmergedBefore = RunGit(absoluteRoot, "ls-files", "--unmerged", "--", gitPath).Output;
        RunGit(absoluteRoot, "add", "--", gitPath);

        var unmergedAfter = RunGit(absoluteRoot, "ls-files", "--unmerged", "--", gitPath).Output;
        if (!string.IsNullOrWhiteSpace(unmergedAfter))
            throw new GitStagingException($"git add 完成后文件仍处于冲突状态：{gitPath}。");
        var stageEntry = RunGit(absoluteRoot, "ls-files", "--stage", "--", gitPath).Output;
        if (!HasStageZeroEntry(stageEntry))
            throw new GitStagingException($"Git index 中没有 MERGED 的 stage-0 条目：{gitPath}。");

        return new GitStageResult(absoluteRoot, gitPath, !string.IsNullOrWhiteSpace(unmergedBefore));
    }

    private static bool IsWithinRoot(string root, string path)
    {
        var relative = Path.GetRelativePath(Path.GetFullPath(root), Path.GetFullPath(path));
        return !Path.IsPathRooted(relative) &&
               !relative.Equals("..", StringComparison.Ordinal) &&
               !relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) &&
               !relative.StartsWith($"..{Path.AltDirectorySeparatorChar}", StringComparison.Ordinal);
    }

    private static bool HasStageZeroEntry(string output) => output
        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
        .Any(line => line.Contains(" 0\t", StringComparison.Ordinal));

    private static GitProcessResult RunGit(string workingDirectory, params string[] arguments) =>
        RunGit(workingDirectory, allowFailure: false, arguments);

    private static GitProcessResult RunGit(
        string workingDirectory,
        bool allowFailure,
        params string[] arguments)
    {
        try
        {
            var startInfo = new ProcessStartInfo("git")
            {
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (var argument in arguments)
                startInfo.ArgumentList.Add(argument);
            using var process = Process.Start(startInfo)
                ?? throw new GitStagingException("无法启动 git 进程。");
            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit((int)GitTimeout.TotalMilliseconds))
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit();
                throw new GitStagingException($"Git 命令执行超过 {GitTimeout.TotalSeconds:0} 秒。");
            }
            Task.WaitAll(outputTask, errorTask);
            var result = new GitProcessResult(process.ExitCode, outputTask.Result, errorTask.Result);
            if (!allowFailure && result.ExitCode != 0)
            {
                var details = string.IsNullOrWhiteSpace(result.Error) ? result.Output : result.Error;
                throw new GitStagingException(
                    $"Git 命令失败（退出码 {result.ExitCode}）：git {string.Join(' ', arguments)}。{Environment.NewLine}{details.Trim()}");
            }
            return result;
        }
        catch (GitStagingException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException or InvalidOperationException)
        {
            throw new GitStagingException("执行 Git 自动暂存失败。", exception);
        }
    }

    private sealed record GitProcessResult(int ExitCode, string Output, string Error);
}
