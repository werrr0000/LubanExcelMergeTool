using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;

namespace LubanExcelMerge.Cli;

public sealed record DiagnosticPackageResult(string OutputPath, string LogSha256);

public static class DiagnosticPackageExporter
{
    private const long MaximumLogBytes = 64L * 1024 * 1024;

    public static DiagnosticPackageResult Export(string logPath, string outputPath)
    {
        var absoluteLog = Path.GetFullPath(logPath);
        var absoluteOutput = Path.GetFullPath(outputPath);
        if (!File.Exists(absoluteLog))
            throw new MergeInputException($"诊断日志不存在：{absoluteLog}。");
        if (!string.Equals(Path.GetExtension(absoluteLog), ".jsonl", StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("诊断日志必须使用 .jsonl 扩展名。");
        if (!string.Equals(Path.GetExtension(absoluteOutput), ".zip", StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("诊断包输出必须使用 .zip 扩展名。");
        if (string.Equals(absoluteLog, absoluteOutput, StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("诊断包不能覆盖诊断日志。");

        var directory = Path.GetDirectoryName(absoluteOutput)
            ?? throw new MergeInputException("诊断包输出目录无效。");
        Directory.CreateDirectory(directory);
        ValidateLog(absoluteLog);
        var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(absoluteOutput)}.{Guid.NewGuid():N}.tmp");
        string hash;
        using (var logStream = File.OpenRead(absoluteLog))
            hash = Convert.ToHexString(SHA256.HashData(logStream));

        try
        {
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Create))
            {
                archive.CreateEntryFromFile(absoluteLog, "diagnostics.jsonl", CompressionLevel.Optimal);
                var manifestEntry = archive.CreateEntry("manifest.json", CompressionLevel.Optimal);
                using var manifestStream = manifestEntry.Open();
                JsonSerializer.Serialize(manifestStream, new
                {
                    formatVersion = 1,
                    createdUtc = DateTimeOffset.UtcNow,
                    toolVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(),
                    operatingSystem = RuntimeInformation.OSDescription,
                    log = new { entry = "diagnostics.jsonl", sha256 = hash },
                    includesWorkbooks = false
                }, new JsonSerializerOptions { WriteIndented = true });
            }

            if (File.Exists(absoluteOutput))
                File.Replace(temporaryPath, absoluteOutput, null, ignoreMetadataErrors: true);
            else
                File.Move(temporaryPath, absoluteOutput);
            return new DiagnosticPackageResult(absoluteOutput, hash);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            throw new WorkbookWriteException("导出诊断包失败。", exception);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }
    }

    private static void ValidateLog(string logPath)
    {
        var length = new FileInfo(logPath).Length;
        if (length == 0 || length > MaximumLogBytes)
            throw new MergeInputException("诊断日志为空或超过 64 MB，无法导出。");

        try
        {
            foreach (var line in File.ReadLines(logPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;
                using var document = JsonDocument.Parse(line);
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                    throw new JsonException("JSON Lines 根节点必须是对象。");
            }
        }
        catch (JsonException exception)
        {
            throw new MergeInputException("诊断日志不是有效的 JSON Lines 文件。", exception);
        }
        catch (IOException exception)
        {
            throw new MergeInputException("无法读取诊断日志。", exception);
        }
    }
}

public static class DiagnosticPackageCommand
{
    public static int Run(IReadOnlyList<string> args, TextWriter output)
    {
        string? logPath = null;
        string? outputPath = null;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 1; index < args.Count; index += 2)
        {
            if (index + 1 >= args.Count)
                throw new MergeInputException($"参数 {args[index]} 缺少值。");
            var option = args[index];
            if (!seen.Add(option))
                throw new MergeInputException($"参数 {option} 重复。");
            switch (option)
            {
                case "--log":
                    logPath = args[index + 1];
                    break;
                case "--output":
                    outputPath = args[index + 1];
                    break;
                default:
                    throw new MergeInputException($"未知参数：{args[index]}。");
            }
        }

        if (string.IsNullOrWhiteSpace(logPath) || string.IsNullOrWhiteSpace(outputPath))
            throw new MergeInputException("diagnostic-package 需要 --log 和 --output。");
        var result = DiagnosticPackageExporter.Export(logPath, outputPath);
        output.WriteLine($"诊断包已导出：{result.OutputPath}");
        output.WriteLine("诊断包不包含原始工作簿。");
        return ExitCodes.Success;
    }
}
