using System.IO.Compression;
using System.Security.Cryptography;

namespace LubanExcelMerge.OpenXml;

internal static class PackageIntegrity
{
    internal static IReadOnlyDictionary<string, string> HashParts(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, HashEntry, StringComparer.Ordinal);
    }

    private static string HashEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
