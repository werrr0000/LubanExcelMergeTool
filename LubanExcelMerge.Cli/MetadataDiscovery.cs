namespace LubanExcelMerge.Cli;

public sealed record MetadataLocation(string DataRoot, string TablesPath);

public static class MetadataDiscovery
{
    public static MetadataLocation Discover(MergeCommandOptions options)
    {
        if (options.TablesPath is not null)
        {
            var tables = RequireFile(options.TablesPath, "Luban 表定义");
            var dataRoot = options.DataRoot is null
                ? Path.GetDirectoryName(tables)!
                : RequireDirectory(options.DataRoot, "Luban 数据目录");
            return new MetadataLocation(dataRoot, tables);
        }

        if (options.DataRoot is not null)
        {
            var dataRoot = RequireDirectory(options.DataRoot, "Luban 数据目录");
            return new MetadataLocation(dataRoot, RequireFile(Path.Combine(dataRoot, "__tables__.csv"), "Luban 表定义"));
        }

        var repositoryRoot = RequireDirectory(options.RepositoryRoot, "Git 仓库根目录");
        var repositoryTables = Path.Combine(repositoryRoot, "ConfigLuban", "Datas", "__tables__.csv");
        if (File.Exists(repositoryTables))
            return new MetadataLocation(Path.GetDirectoryName(repositoryTables)!, Path.GetFullPath(repositoryTables));

        foreach (var path in new[] { options.OutputPath, options.LocalPath, options.RemotePath })
        {
            var current = new DirectoryInfo(Path.GetDirectoryName(Path.GetFullPath(path))!);
            while (current is not null)
            {
                var directTables = Path.Combine(current.FullName, "__tables__.csv");
                if (File.Exists(directTables) && string.Equals(current.Name, "Datas", StringComparison.OrdinalIgnoreCase))
                    return new MetadataLocation(current.FullName, directTables);

                var nestedTables = Path.Combine(current.FullName, "ConfigLuban", "Datas", "__tables__.csv");
                if (File.Exists(nestedTables))
                    return new MetadataLocation(Path.GetDirectoryName(nestedTables)!, nestedTables);
                current = current.Parent;
            }
        }

        throw new UnsafeWorkbookException("未找到 ConfigLuban/Datas/__tables__.csv，无法在无界面模式下确定稳定主键。");
    }

    private static string RequireFile(string path, string description)
    {
        var absolute = Path.GetFullPath(path);
        if (!File.Exists(absolute))
            throw new MergeInputException($"{description}不存在：{absolute}。");
        return absolute;
    }

    private static string RequireDirectory(string path, string description)
    {
        var absolute = Path.GetFullPath(path);
        if (!Directory.Exists(absolute))
            throw new MergeInputException($"{description}不存在：{absolute}。");
        return absolute;
    }
}
