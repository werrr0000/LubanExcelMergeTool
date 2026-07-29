namespace LubanExcelMerge.Cli;

public interface IFullExportValidator
{
    void Validate(string commandPath, string repositoryRoot, TimeSpan timeout);
}

public sealed class IsolatedFullExportValidator : IFullExportValidator
{
    private const int MaximumFiles = 200_000;
    private const long MaximumBytes = 4L * 1024 * 1024 * 1024;
    private readonly IProjectValidator _commandRunner;

    public IsolatedFullExportValidator(IProjectValidator? commandRunner = null)
    {
        _commandRunner = commandRunner ?? new ProjectValidationRunner("完整导出校验");
    }

    public void Validate(string commandPath, string repositoryRoot, TimeSpan timeout)
    {
        var sourceRoot = Path.GetFullPath(repositoryRoot);
        var sourceCommand = Path.GetFullPath(commandPath);
        var commandRelativePath = Path.GetRelativePath(sourceRoot, sourceCommand);
        if (IsOutsideRoot(commandRelativePath))
            throw new ProjectValidationException("完整导出校验命令必须位于仓库目录内，才能在隔离副本中安全运行。");

        var sourceConfig = Path.Combine(sourceRoot, "ConfigLuban");
        if (!Directory.Exists(sourceConfig))
            throw new ProjectValidationException($"完整导出校验缺少目录：{sourceConfig}。");

        var isolationRoot = Path.Combine(
            Path.GetTempPath(),
            $"LubanExcelMerge.FullExport-{Guid.NewGuid():N}");
        var isolatedRepository = Path.Combine(isolationRoot, "Repository");
        try
        {
            Directory.CreateDirectory(isolatedRepository);
            var budget = new CopyBudget();
            CopyDirectory(sourceConfig, Path.Combine(isolatedRepository, "ConfigLuban"), budget);
            var sourceOutput = Path.Combine(sourceRoot, "ConfigOutput");
            if (Directory.Exists(sourceOutput))
                CopyDirectory(sourceOutput, Path.Combine(isolatedRepository, "ConfigOutput"), budget);

            var isolatedCommand = Path.Combine(isolatedRepository, commandRelativePath);
            if (!File.Exists(isolatedCommand))
                throw new ProjectValidationException($"隔离副本中没有完整导出校验命令：{isolatedCommand}。");
            _commandRunner.Validate(isolatedCommand, isolatedRepository, timeout);
        }
        catch (ProjectValidationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new ProjectValidationException("创建或运行完整导出隔离副本失败。", exception);
        }
        finally
        {
            TryDeleteIsolationRoot(isolationRoot);
        }
    }

    private static void CopyDirectory(string source, string target, CopyBudget budget)
    {
        var pending = new Stack<(DirectoryInfo Source, string Target)>();
        pending.Push((new DirectoryInfo(source), target));
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            if ((current.Source.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new ProjectValidationException($"完整导出隔离副本不允许目录链接：{current.Source.FullName}。");
            Directory.CreateDirectory(current.Target);

            foreach (var file in current.Source.EnumerateFiles())
            {
                if ((file.Attributes & FileAttributes.ReparsePoint) != 0)
                    throw new ProjectValidationException($"完整导出隔离副本不允许文件链接：{file.FullName}。");
                budget.Add(file.Length);
                File.Copy(file.FullName, Path.Combine(current.Target, file.Name), overwrite: false);
            }

            foreach (var directory in current.Source.EnumerateDirectories())
            {
                if (directory.Name is ".git" or "temp")
                    continue;
                pending.Push((directory, Path.Combine(current.Target, directory.Name)));
            }
        }
    }

    private static bool IsOutsideRoot(string relativePath) =>
        Path.IsPathRooted(relativePath) ||
        relativePath.Equals("..", StringComparison.Ordinal) ||
        relativePath.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal);

    private static void TryDeleteIsolationRoot(string isolationRoot)
    {
        try
        {
            if (Directory.Exists(isolationRoot))
                Directory.Delete(isolationRoot, recursive: true);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            // A leftover isolated copy is safer than mutating the source repository during cleanup recovery.
        }
    }

    private sealed class CopyBudget
    {
        private int _files;
        private long _bytes;

        public void Add(long length)
        {
            _files++;
            _bytes += length;
            if (_files > MaximumFiles || _bytes > MaximumBytes)
            {
                throw new ProjectValidationException(
                    $"完整导出隔离副本超过安全上限（{MaximumFiles} 个文件或 {MaximumBytes / 1024 / 1024} MB）。");
            }
        }
    }
}
