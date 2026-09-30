namespace LubanExcelMerge.Git;

public static class GitPath
{
    // Git 返回物理路径；macOS 的 /var、/tmp 等入口则可能包含父级符号链接。
    public static string GetPhysicalPath(string path)
    {
        var absolute = Path.GetFullPath(path);
        var current = Path.GetPathRoot(absolute)!;
        foreach (var part in absolute[current.Length..].Split(Path.DirectorySeparatorChar,
                     StringSplitOptions.RemoveEmptyEntries))
        {
            current = Path.Combine(current, part);
            FileSystemInfo info = Directory.Exists(current) ? new DirectoryInfo(current) : new FileInfo(current);
            if (info.LinkTarget is not null)
                current = info.ResolveLinkTarget(returnFinalTarget: true)?.FullName
                    ?? throw new IOException($"无法解析符号链接：{current}");
        }
        return current;
    }
}
