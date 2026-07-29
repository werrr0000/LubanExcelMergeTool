using System.Security.Cryptography;
using System.Text;

namespace LubanExcelMerge.Cli;

public sealed class MergeOutputLease : IDisposable
{
    private readonly Mutex _mutex;
    private bool _ownsMutex;

    private MergeOutputLease(Mutex mutex, bool ownsMutex)
    {
        _mutex = mutex;
        _ownsMutex = ownsMutex;
    }

    public static MergeOutputLease Acquire(string outputPath)
    {
        var normalized = Path.GetFullPath(outputPath).ToUpperInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(normalized)));
        var mutex = new Mutex(false, $"LubanExcelMerge-{hash}");
        try
        {
            var ownsMutex = false;
            try
            {
                ownsMutex = mutex.WaitOne(TimeSpan.Zero);
            }
            catch (AbandonedMutexException)
            {
                ownsMutex = true;
            }

            if (!ownsMutex)
                throw new WorkbookWriteException("另一个合并进程正在写入同一个 MERGED 文件，请稍后重试。");
            return new MergeOutputLease(mutex, ownsMutex: true);
        }
        catch
        {
            mutex.Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        if (_ownsMutex)
        {
            _mutex.ReleaseMutex();
            _ownsMutex = false;
        }
        _mutex.Dispose();
    }
}

public static class MergeOutputRecovery
{
    private const string BackupMarker = ".validation-backup-";
    private const string NewOutputMarker = ".validation-new-output-";

    public static string CreateRollbackMarker(string outputPath, bool outputExisted)
    {
        var absolute = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(absolute)
            ?? throw new WorkbookWriteException("MERGED 输出目录无效。");
        Directory.CreateDirectory(directory);
        var marker = Path.Combine(
            directory,
            $".{Path.GetFileName(absolute)}{(outputExisted ? BackupMarker : NewOutputMarker)}{Guid.NewGuid():N}");
        if (outputExisted)
            File.Copy(absolute, marker, overwrite: false);
        else
            using (File.Create(marker)) { }
        return marker;
    }

    public static void RecoverPending(string outputPath)
    {
        var absolute = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(absolute)
            ?? throw new WorkbookWriteException("MERGED 输出目录无效。");
        if (!Directory.Exists(directory))
            return;

        var fileName = Path.GetFileName(absolute);
        var backups = Directory.EnumerateFiles(directory, $".{fileName}{BackupMarker}*").ToArray();
        var newOutputMarkers = Directory.EnumerateFiles(directory, $".{fileName}{NewOutputMarker}*").ToArray();
        if (backups.Length + newOutputMarkers.Length > 1)
        {
            throw new WorkbookWriteException(
                $"检测到多个未完成的 MERGED 保存事务，已保留诊断文件，请人工确认：{absolute}。");
        }

        try
        {
            if (backups.Length == 1)
            {
                File.Copy(backups[0], absolute, overwrite: true);
                File.Delete(backups[0]);
            }
            else if (newOutputMarkers.Length == 1)
            {
                if (File.Exists(absolute))
                    File.Delete(absolute);
                File.Delete(newOutputMarkers[0]);
            }

            foreach (var temporaryPath in Directory.EnumerateFiles(directory, $".{fileName}.*.tmp.xlsx"))
                File.Delete(temporaryPath);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new WorkbookWriteException("无法恢复上次未完成的 MERGED 保存事务。", exception);
        }
    }

    public static void Restore(string outputPath, bool outputExisted, string markerPath)
    {
        try
        {
            if (outputExisted)
                File.Copy(markerPath, outputPath, overwrite: true);
            else if (File.Exists(outputPath))
                File.Delete(outputPath);
            File.Delete(markerPath);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new WorkbookWriteException(
                "校验失败，并且无法恢复校验前的 MERGED 文件；恢复标记已保留。",
                exception);
        }
    }

    public static void Commit(string markerPath)
    {
        try
        {
            File.Delete(markerPath);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new WorkbookWriteException("MERGED 已通过校验，但无法提交保存事务。", exception);
        }
    }
}
