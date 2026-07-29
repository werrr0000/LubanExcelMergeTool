using System.Text.RegularExpressions;

namespace LubanExcelMerge.Cli;

public static class PathPatternMatcher
{
    public static bool IsMatch(string repositoryRelativePath, IReadOnlyList<string>? patterns)
    {
        if (patterns is null || patterns.Count == 0)
            return false;
        var path = NormalizePath(repositoryRelativePath);
        return patterns.Any(pattern => CreateRegex(NormalizePattern(pattern)).IsMatch(path));
    }

    public static string NormalizePattern(string pattern) => NormalizePath(pattern.Trim());

    private static string NormalizePath(string path) => path.Replace('\\', '/').TrimStart('/');

    private static Regex CreateRegex(string pattern)
    {
        var expression = Regex.Escape(pattern)
            .Replace("\\*\\*", ".*", StringComparison.Ordinal)
            .Replace("\\*", "[^/]*", StringComparison.Ordinal)
            .Replace("\\?", "[^/]", StringComparison.Ordinal);
        return new Regex($"^{expression}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    }
}
