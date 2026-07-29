using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using LubanExcelMerge.Git;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Cli;

public sealed class MergeDiagnosticLogger
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly string _path;
    private readonly object _writeLock = new();

    private MergeDiagnosticLogger(string path)
    {
        _path = path;
    }

    public string Path => _path;

    public static MergeDiagnosticLogger? Create(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;
        var absolute = System.IO.Path.GetFullPath(path);
        try
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(absolute)!);
            using var stream = new FileStream(absolute, FileMode.Append, FileAccess.Write, FileShare.Read);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new MergeInputException($"无法创建诊断日志：{absolute}。", exception);
        }
        return new MergeDiagnosticLogger(absolute);
    }

    public static string CreateDefaultPath()
    {
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return System.IO.Path.Combine(
            root,
            "LubanExcelMerge",
            "logs",
            DateTime.Now.ToString("yyyy-MM-dd"),
            $"LubanExcelMerge-{DateTime.Now:HHmmss}-{Guid.NewGuid():N}.jsonl");
    }

    public void WriteStarted(MergeCommandOptions options)
    {
        Write("started", new
        {
            toolVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(),
            operatingSystem = RuntimeInformation.OSDescription,
            processArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
            mode = options.Headless ? "headless" : "interactive",
            repositoryRoot = System.IO.Path.GetFullPath(options.RepositoryRoot),
            files = new
            {
                @base = GetFileIdentity(options.BasePath),
                local = GetFileIdentity(options.LocalPath),
                remote = GetFileIdentity(options.RemotePath),
                merged = GetFileIdentity(options.OutputPath)
            }
        });
    }

    public void WritePrepared(PreparedMergeSession session)
    {
        Write("prepared", new
        {
            logicalTable = session.LogicalTable,
            sheet = session.SheetName,
            key = session.KeyName,
            automaticEdits = session.AutomaticEditCount,
            conflicts = session.Conflicts.Count,
            changedCells = session.ChangedCells,
            addedRecords = session.AddedRecords,
            deletedRecords = session.DeletedRecords,
            metadataChanges = session.MetadataChangeCount,
            ignoredFields = session.IgnoredFields,
            preparationMilliseconds = new
            {
                total = session.PreparationTimings.TotalMilliseconds,
                workbookRead = session.PreparationTimings.WorkbookReadMilliseconds,
                @base = session.PreparationTimings.BaseWorkbookReadMilliseconds,
                local = session.PreparationTimings.LocalWorkbookReadMilliseconds,
                remote = session.PreparationTimings.RemoteWorkbookReadMilliseconds,
                sheets = session.PreparationTimings.SheetPreparationMilliseconds
            },
            sheets = session.Sheets.Select(sheet => new
            {
                name = sheet.SheetName,
                key = sheet.KeyName,
                automaticEdits = sheet.AutomaticEditCount,
                conflicts = sheet.Conflicts.Count,
                remainingConflicts = sheet.RemainingConflicts,
                changedCells = sheet.ChangedCells,
                addedRecords = sheet.AddedRecords,
                deletedRecords = sheet.DeletedRecords,
                metadataChanges = sheet.MetadataChangeCount
            }),
            logicalTableUniquenessValidated = session.LogicalTableUniquenessValidated
        });
    }

    public void WriteCompleted(MergeRunResult result)
    {
        Write("completed", new
        {
            result.Succeeded,
            logicalTable = result.LogicalTable,
            sheet = result.SheetName,
            key = result.KeyName,
            conflicts = result.Conflicts.Count,
            changedCells = result.ChangedCells,
            addedRecords = result.AddedRecords,
            deletedRecords = result.DeletedRecords,
            formulaStatus = result.RecalculationStatus?.ToString(),
            formulaProvider = result.RecalculationProvider,
            projectValidationCompleted = result.ProjectValidationCompleted,
            fullExportValidationCompleted = result.FullExportValidationCompleted,
            logicalTableUniquenessValidated = result.LogicalTableUniquenessValidated,
            ignoredFields = result.IgnoredFields,
            preparationMilliseconds = new
            {
                total = result.PreparationTimings.TotalMilliseconds,
                workbookRead = result.PreparationTimings.WorkbookReadMilliseconds,
                @base = result.PreparationTimings.BaseWorkbookReadMilliseconds,
                local = result.PreparationTimings.LocalWorkbookReadMilliseconds,
                remote = result.PreparationTimings.RemoteWorkbookReadMilliseconds,
                sheets = result.PreparationTimings.SheetPreparationMilliseconds
            },
            merged = result.OutputPath is null ? null : GetFileIdentity(result.OutputPath)
        });
    }

    public void WriteSaved(PreparedMergeSession session, WorkbookSaveResult result, bool gitStaged = false)
    {
        Write("completed", new
        {
            succeeded = true,
            logicalTable = session.LogicalTable,
            sheet = session.SheetName,
            key = session.KeyName,
            conflicts = 0,
            changedCells = session.ChangedCells,
            addedRecords = session.AddedRecords,
            deletedRecords = session.DeletedRecords,
            formulaCount = result.FormulaCount,
            formulaStatus = result.RecalculationStatus.ToString(),
            formulaProvider = result.RecalculationProvider,
            projectValidationCompleted = result.ProjectValidationCompleted,
            fullExportValidationCompleted = result.FullExportValidationCompleted,
            gitStaged,
            logicalTableUniquenessValidated = session.LogicalTableUniquenessValidated,
            ignoredFields = session.IgnoredFields,
            metadataChanges = session.MetadataChangeCount,
            sheets = session.Sheets.Select(sheet => new
            {
                name = sheet.SheetName,
                key = sheet.KeyName,
                automaticEdits = sheet.AutomaticEditCount,
                conflicts = sheet.Conflicts.Count,
                remainingConflicts = sheet.RemainingConflicts,
                changedCells = sheet.ChangedCells,
                addedRecords = sheet.AddedRecords,
                deletedRecords = sheet.DeletedRecords,
                metadataChanges = sheet.MetadataChangeCount
            }),
            merged = GetFileIdentity(result.OutputPath)
        });
    }

    public void WriteException(Exception exception, int exitCode)
    {
        Write("exception", new
        {
            exitCode,
            exceptionType = exception.GetType().FullName,
            innerExceptionType = exception.InnerException?.GetType().FullName,
            stackTrace = exception.StackTrace
        });
    }

    private void Write(string eventName, object details)
    {
        try
        {
            var line = JsonSerializer.Serialize(new
            {
                timestampUtc = DateTimeOffset.UtcNow,
                @event = eventName,
                details
            }, JsonOptions);
            lock (_writeLock)
                File.AppendAllText(_path, line + Environment.NewLine, new UTF8Encoding(false));
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            // Logging must never turn a validated MERGED result into a failed merge.
        }
    }

    private static object GetFileIdentity(string path)
    {
        var absolute = System.IO.Path.GetFullPath(path);
        if (!File.Exists(absolute))
            return new { path = absolute, exists = false, sha256 = (string?)null, isLfsPointer = false };
        using var stream = File.OpenRead(absolute);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var isLfsPointer = false;
        try
        {
            isLfsPointer = GitLfsPointer.TryRead(absolute) is not null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException)
        {
            // The hash and path remain useful even when pointer probing fails.
        }
        return new { path = absolute, exists = true, sha256 = hash, isLfsPointer };
    }
}
