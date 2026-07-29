using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace LubanExcelMerge.Git;

public sealed record GitLfsPointer(string Sha256, long Size)
{
    private static readonly Regex ObjectIdPattern = new(
        "^[0-9a-f]{64}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static GitLfsPointer? TryRead(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > 2_048)
            return null;

        var prefix = new byte[8];
        if (stream.Read(prefix, 0, prefix.Length) != prefix.Length ||
            !prefix.AsSpan().SequenceEqual("version "u8))
            return null;
        stream.Position = 0;

        using var reader = new StreamReader(
            stream,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
            detectEncodingFromByteOrderMarks: true);
        var text = reader.ReadToEnd().Replace("\r\n", "\n", StringComparison.Ordinal);
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length != 3 || lines[0] != "version https://git-lfs.github.com/spec/v1")
            return null;
        if (!lines[1].StartsWith("oid sha256:", StringComparison.Ordinal) ||
            !lines[2].StartsWith("size ", StringComparison.Ordinal))
            return null;

        var objectId = lines[1]["oid sha256:".Length..];
        if (!ObjectIdPattern.IsMatch(objectId) ||
            !long.TryParse(lines[2]["size ".Length..], NumberStyles.None, CultureInfo.InvariantCulture, out var size) ||
            size < 0)
            throw new InvalidDataException($"文件 {path} 是损坏的 Git LFS 指针。");

        return new GitLfsPointer(objectId, size);
    }
}
