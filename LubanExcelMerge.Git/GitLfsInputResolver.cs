using System.Security.Cryptography;

namespace LubanExcelMerge.Git;

public sealed record ResolvedInput(string OriginalPath, string ContentPath, bool IsLfsPointer);

public sealed class LfsObjectNotFoundException : IOException
{
    public LfsObjectNotFoundException(string message) : base(message)
    {
    }
}

public sealed class GitLfsInputResolver
{
    public ResolvedInput Resolve(string inputPath, string repositoryRoot)
    {
        var absoluteInput = Path.GetFullPath(inputPath);
        if (new FileInfo(absoluteInput).Length <= 4096)
        {
            var prefix = System.Text.Encoding.UTF8.GetString(File.ReadAllBytes(absoluteInput));
            if (prefix.StartsWith("version https://git-lfs.github.com/spec/v1", StringComparison.Ordinal) &&
                prefix.Contains("<<<<<<<", StringComparison.Ordinal) && prefix.Contains(">>>>>>>", StringComparison.Ordinal))
                throw new InvalidDataException("输入文件是尚未解决的 Git LFS 冲突文本，不是 Excel 工作簿。请使用 Git/Fork 提供的 BASE、LOCAL、REMOTE 三方文件进行合并，完成保存前不能用 Excel 打开工作区的冲突文件。");
        }
        var pointer = GitLfsPointer.TryRead(absoluteInput);
        if (pointer is null)
            return new ResolvedInput(absoluteInput, absoluteInput, false);

        var gitDirectory = ResolveCommonGitDirectory(Path.GetFullPath(repositoryRoot));
        var objectPath = Path.Combine(
            gitDirectory,
            "lfs",
            "objects",
            pointer.Sha256[..2],
            pointer.Sha256.Substring(2, 2),
            pointer.Sha256);
        if (!File.Exists(objectPath))
            throw new LfsObjectNotFoundException(
                $"Git LFS 对象 {pointer.Sha256} 不在本地。请先在仓库中执行 git lfs fetch。");

        var info = new FileInfo(objectPath);
        if (info.Length != pointer.Size)
            throw new InvalidDataException($"Git LFS 对象 {pointer.Sha256} 的大小与指针不一致。");
        using var stream = File.OpenRead(objectPath);
        var actualHash = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        if (!string.Equals(actualHash, pointer.Sha256, StringComparison.Ordinal))
            throw new InvalidDataException($"Git LFS 对象 {pointer.Sha256} 的 SHA-256 校验失败。");

        return new ResolvedInput(absoluteInput, objectPath, true);
    }

    private static string ResolveCommonGitDirectory(string repositoryRoot)
    {
        var dotGit = Path.Combine(repositoryRoot, ".git");
        string gitDirectory;
        if (Directory.Exists(dotGit))
        {
            gitDirectory = dotGit;
        }
        else if (File.Exists(dotGit))
        {
            var text = File.ReadAllText(dotGit).Trim();
            const string prefix = "gitdir:";
            if (!text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"{dotGit} 不是有效的 gitdir 文件。");
            var value = text[prefix.Length..].Trim();
            gitDirectory = Path.GetFullPath(value, repositoryRoot);
        }
        else
        {
            throw new DirectoryNotFoundException($"仓库 {repositoryRoot} 中不存在 .git。");
        }

        var commonDirectoryFile = Path.Combine(gitDirectory, "commondir");
        if (!File.Exists(commonDirectoryFile))
            return gitDirectory;

        var commonDirectory = File.ReadAllText(commonDirectoryFile).Trim();
        return Path.GetFullPath(commonDirectory, gitDirectory);
    }
}
