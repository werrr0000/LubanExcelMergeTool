using System.IO.Compression;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Git;
using LubanExcelMerge.OpenXml;
using Core = LubanExcelMerge.Core;

if (args is ["--create-ui-fixture", var fixtureRoot])
{
    var fixture = CreateScenario(fixtureRoot);
    TestWorkbookFactory.Create(fixture.BasePath, new[]
    {
        new[] { "1001", "旧名称", "基础说明" },
        new[] { "1002", "待删除", "保持不变" },
        new[] { "1003", "同键新增基础", "基础" }
    });
    TestWorkbookFactory.Create(fixture.LocalPath, new[]
    {
        new[] { "1001", "本地名称", "基础说明" },
        new[] { "1002", "本地已修改", "保持不变" },
        new[] { "2000", "本地新增", "LOCAL" }
    });
    TestWorkbookFactory.Create(fixture.RemotePath, new[]
    {
        new[] { "1001", "远端名称", "远端说明" },
        new[] { "1003", "同键新增基础", "基础" },
        new[] { "2000", "远端新增", "REMOTE" },
        new[] { "3000", "自动新增", "REMOTE" }
    });
    Console.WriteLine(string.Join(Environment.NewLine, CreateArguments(fixture)));
    return 0;
}

    if (args is ["--create-multi-sheet-ui-fixture", var multiSheetFixtureRoot])
{
    var fixture = CreateScenario(multiSheetFixtureRoot);
    TestWorkbookFactory.CreateMultiSheet(
        fixture.BasePath,
        ("角色属性", new[] { new[] { "1001", "基础攻击", "10" }, new[] { "1002", "基础防御", "20" } }),
        ("关卡参数", new[] { new[] { "2001", "普通难度", "1" }, new[] { "2002", "困难难度", "2" } }));
    TestWorkbookFactory.CreateMultiSheet(
        fixture.LocalPath,
        ("角色属性", new[] { new[] { "1001", "本地攻击", "10" }, new[] { "1002", "基础防御", "20" } }),
        ("关卡参数", new[] { new[] { "2001", "本地普通", "1" }, new[] { "2002", "困难难度", "2" } }));
    TestWorkbookFactory.CreateMultiSheet(
        fixture.RemotePath,
        ("角色属性", new[] { new[] { "1001", "远端攻击", "11" }, new[] { "1002", "基础防御", "20" } }),
        ("关卡参数", new[] { new[] { "2001", "远端普通", "1" }, new[] { "2002", "困难难度", "3" } }));
    Console.WriteLine(string.Join(Environment.NewLine, CreateArguments(fixture)));
    return 0;
}

if (args is ["--metadata-row-tests"])
{
    var focusedRoot = Path.Combine(Path.GetTempPath(), "LubanExcelMerge.Cli.Tests", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(focusedRoot);
    try
    {
        InsertedMetadataRowRequiresReview(focusedRoot);
        DeletedMetadataRowRequiresReview(focusedRoot);
        Console.WriteLine("PASS metadata row insertion and deletion");
        return 0;
    }
    finally
    {
        if (Directory.Exists(focusedRoot))
            Directory.Delete(focusedRoot, recursive: true);
    }
}

var testRoot = Path.Combine(Path.GetTempPath(), "LubanExcelMerge.Cli.Tests", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testRoot);
try
{
    var tests = new List<(string Name, Action Run)>
    {
        ("command line requires four files and repository", CommandLineRequiresInputs),
        ("command line accepts Fork arguments without merge subcommand", CommandLineAcceptsForkArgumentsWithoutSubcommand),
        ("command line still rejects unknown subcommands", CommandLineRejectsUnknownSubcommands),
        ("Git LFS pointer resolves verified local object", () => LfsPointerResolves(testRoot)),
        ("missing Git LFS object gives actionable error", () => MissingLfsObject(testRoot)),
        ("small binary file is not decoded as LFS text", () => SmallBinaryIsNotLfs(testRoot)),
        ("CLI maps missing LFS object to exit 2", () => CliMapsMissingLfsObject(testRoot)),
        ("Git config command emits four-file protocol without mutation", () => GitConfigCommandDoesNotMutate(testRoot)),
        ("direct GUI-style save resolves and stages only MERGED", () => GitStagerResolvesAndStagesOnlyMerged(testRoot)),
        ("stager discovers nested ConfigLuban repository", () => GitStagerDiscoversNestedConfigRepository(testRoot)),
        ("real Git LFS mergetool completes four-file protocol", () => GitLfsMergetoolEndToEnd(testRoot)),
        ("non-conflicting four-file merge succeeds", () => NonConflictingMerge(testRoot)),
        ("mode one merges the single record by field", () => ModeOneMergesSingleRecord(testRoot)),
        ("mode one field conflict requires resolution", () => ModeOneFieldConflictRequiresResolution(testRoot)),
        ("mode one rejects multiple data records", () => ModeOneRejectsMultipleRecords(testRoot)),
        ("automatic merge results include LOCAL and REMOTE changes", () => AutomaticMergeResultsIncludeBothSides(testRoot)),
        ("serialized blank rows and columns do not displace additions", () => BlankLayoutDoesNotDisplaceAdditions(testRoot)),
        ("large table preparation and navigation stay bounded", () => LargeTablePerformanceSmoke(testRoot)),
        ("multi-sheet automatic edits save every sheet", () => MultiSheetAutomaticMerge(testRoot)),
        ("multi-sheet save requires every conflict resolved", () => MultiSheetRequiresEveryConflictResolved(testRoot)),
        ("multi-sheet names and order must match", () => MultiSheetNamesMustMatch(testRoot)),
        ("metadata change requires review and saves selected value", () => MetadataChangeRequiresReview(testRoot)),
        ("inserted metadata row requires review and shifts data", () => InsertedMetadataRowRequiresReview(testRoot)),
        ("deleted metadata row requires review and shifts data", () => DeletedMetadataRowRequiresReview(testRoot)),
        ("multi-sheet metadata reviews gate whole workbook save", () => MultiSheetMetadataChangesGateSave(testRoot)),
        ("multi-sheet REMOTE appended fields save every sheet", () => MultiSheetRemoteAppendedFieldsSaveEverySheet(testRoot)),
        ("REMOTE-only appended field merges metadata and data", () => RemoteOnlyAppendedFieldMerges(testRoot)),
        ("different LOCAL and REMOTE appended fields form a union", () => DifferentAppendedFieldsFormUnion(testRoot)),
        ("same appended field data conflict is resolvable", () => SameAppendedFieldDataConflictIsResolvable(testRoot)),
        ("same appended field type conflict requires review", () => SameAppendedFieldTypeConflictRequiresReview(testRoot)),
        ("REMOTE field inserted inside layout reorders existing fields", () => RemoteInsertedFieldReordersExistingFields(testRoot)),
        ("REMOTE inserted field moves and updates formula field", () => RemoteInsertedFieldMovesFormulaField(testRoot)),
        ("LOCAL field inserted inside layout remains the MERGED layout", () => LocalInsertedFieldDefinesMergedLayout(testRoot)),
        ("REMOTE reordered existing fields merge", () => RemoteReorderedExistingFieldsMerge(testRoot)),
        ("REMOTE deleted existing field merges", () => RemoteDeletedExistingFieldMerges(testRoot)),
        ("REMOTE existing field type change merges", () => RemoteExistingFieldTypeChangeMerges(testRoot)),
        ("REMOTE existing field rename merges", () => RemoteExistingFieldRenameMerges(testRoot)),
        ("LOCAL primary-key field rename preserves record identity", () => LocalPrimaryKeyFieldRenamePreservesIdentity(testRoot)),
        ("column delete-modify conflict is resolvable", () => ColumnDeleteModifyConflictIsResolvable(testRoot)),
        ("column modify-modify conflict is resolvable", () => ColumnModifyModifyConflictIsResolvable(testRoot)),
        ("column move-move conflict is resolvable", () => ColumnMoveMoveConflictIsResolvable(testRoot)),
        ("commented data rows do not participate in keys", () => CommentedDataRowsAreIgnored(testRoot)),
        ("conflict returns 1 and preserves MERGED", () => ConflictPreservesMerged(testRoot)),
        ("interactive session resolves one field and keeps automatic fields", () => InteractiveCellResolution(testRoot)),
        ("interactive delete-modify conflict supports remote record", () => InteractiveDeleteModifyResolution(testRoot)),
        ("comparison grid aligns additions and deletions", () => ComparisonGridAlignsRows(testRoot)),
        ("formula recalculation mode reaches atomic save", () => FormulaRecalculationModeReachesSave(testRoot)),
        ("camel-case config drives formula recalculation", () => ConfigDrivesFormulaRecalculation(testRoot)),
        ("strict config rejects unknown fields", () => StrictConfigRejectsUnknownFields(testRoot)),
        ("key override changes record identity", () => KeyOverrideChangesRecordIdentity(testRoot)),
        ("ignored fields preserve LOCAL without conflicts", () => IgnoredFieldsPreserveLocal(testRoot)),
        ("ignored-only modification does not block deletion", () => IgnoredOnlyModificationDoesNotBlockDeletion(testRoot)),
        ("logical-table uniqueness accepts distinct sibling keys", () => LogicalTableUniquenessAcceptsDistinctSiblingKeys(testRoot)),
        ("logical-table uniqueness reports REMOTE sibling collision", () => LogicalTableUniquenessReportsRemoteCollision(testRoot)),
        ("inactive sibling paths are excluded from uniqueness scan", () => InactiveSiblingPathsAreExcluded(testRoot)),
        ("inactive current workbook is rejected", () => InactiveCurrentWorkbookIsRejected(testRoot)),
        ("project validation success is reported", () => ProjectValidationSuccessIsReported(testRoot)),
        ("project validation failure restores MERGED", () => ProjectValidationFailureRestoresMerged(testRoot)),
        ("project validation runner handles batch exit codes", () => ProjectValidationRunnerHandlesBatchExitCodes(testRoot)),
        ("full export runs in isolation without source mutations", () => FullExportRunsInIsolation(testRoot)),
        ("full export failure restores MERGED", () => FullExportFailureRestoresMerged(testRoot)),
        ("pending validation backup restores existing MERGED", () => PendingValidationBackupRestoresExistingMerged(testRoot)),
        ("pending validation marker removes newly-created MERGED", () => PendingValidationMarkerRemovesNewMerged(testRoot)),
        ("diagnostic log records hashes without cell contents", () => DiagnosticLogOmitsCellContents(testRoot)),
        ("diagnostic log records sanitized exception stack", () => DiagnosticLogRecordsSanitizedException(testRoot)),
        ("diagnostic package excludes workbooks", () => DiagnosticPackageExcludesWorkbooks(testRoot)),
        ("diagnostic package rejects disguised workbook", () => DiagnosticPackageRejectsDisguisedWorkbook(testRoot)),
        ("exception types map to documented exit codes", ExceptionTypesMapToExitCodes),
        ("divergent composite-key edits conflict", () => DivergentCompositeKeyEditsConflict(testRoot)),
        ("duplicate key returns unsafe-workbook code", () => DuplicateKeyIsRejected(testRoot))
    };
    if (args is ["--real", var accountPath, var tablesPath, var battlePath])
    {
        tests.Add(("real AccountLv three-way CLI smoke test", () => RealAccountMerge(testRoot, accountPath, tablesPath)));
        tests.Add(("real composite-key CLI smoke test", () => RealBattleMerge(testRoot, battlePath, tablesPath)));
    }

    var failures = new List<string>();
    foreach (var test in tests)
    {
        try
        {
            test.Run();
            Console.WriteLine($"PASS {test.Name}");
        }
        catch (Exception exception)
        {
            failures.Add($"FAIL {test.Name}: {exception}");
        }
    }

    foreach (var failure in failures)
        Console.Error.WriteLine(failure);
    Console.WriteLine($"Executed {tests.Count} tests: {tests.Count - failures.Count} passed, {failures.Count} failed.");
    return failures.Count == 0 ? 0 : 1;
}
finally
{
    var resolvedRoot = Path.GetFullPath(testRoot);
    var resolvedTemp = Path.GetFullPath(Path.GetTempPath());
    if (resolvedRoot.StartsWith(resolvedTemp, StringComparison.OrdinalIgnoreCase) && Directory.Exists(resolvedRoot))
    {
        foreach (var path in Directory.EnumerateFiles(resolvedRoot, "*", SearchOption.AllDirectories))
            File.SetAttributes(path, FileAttributes.Normal);
        Directory.Delete(resolvedRoot, recursive: true);
    }
}

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"Expected {expected}, got {actual}.");
}

static void True(bool value)
{
    if (!value)
        throw new InvalidOperationException("Expected true, got false.");
}

static void Throws<TException>(Action action) where TException : Exception
{
    try
    {
        action();
    }
    catch (TException)
    {
        return;
    }

    throw new InvalidOperationException($"Expected {typeof(TException).Name}.");
}

static string HashFile(string path)
{
    using var stream = File.OpenRead(path);
    return Convert.ToHexString(SHA256.HashData(stream));
}

static void CommandLineRequiresInputs()
{
    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(new[] { "merge", "--base", "base.xlsx" }, output, error);
    Equal(ExitCodes.InvalidInput, exitCode);
    True(error.ToString().Contains("--local", StringComparison.Ordinal));
}

static void CommandLineAcceptsForkArgumentsWithoutSubcommand()
{
    var arguments = new[]
    {
        "--base", "base.xlsx",
        "--local", "local.xlsx",
        "--remote", "remote.xlsx",
        "--output", "merged.xlsx",
        "--repo-root", "repository"
    };

    var options = CommandLineParser.Parse(arguments);

    Equal("base.xlsx", options.BasePath);
    Equal("local.xlsx", options.LocalPath);
    Equal("remote.xlsx", options.RemotePath);
    Equal("merged.xlsx", options.OutputPath);
    Equal("repository", options.RepositoryRoot);
}

static void CommandLineRejectsUnknownSubcommands()
{
    try
    {
        CommandLineParser.Parse(new[] { "unknown", "--base", "base.xlsx" });
        throw new InvalidOperationException("Expected an unknown subcommand error.");
    }
    catch (MergeInputException exception)
    {
        True(exception.Message.Contains("unknown", StringComparison.Ordinal));
    }
}

static void GitConfigCommandDoesNotMutate(string testRoot)
{
    var repository = CreateRepository(Path.Combine(testRoot, "fork config repo"));
    var guiPath = Path.Combine(testRoot, "portable app", "LubanExcelMerge.Gui.exe");
    Directory.CreateDirectory(Path.GetDirectoryName(guiPath)!);
    File.WriteAllText(guiPath, "fixture");
    var gitConfigPath = Path.Combine(repository, ".git", "config");
    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(
        new[] { "git-config", "--gui", guiPath, "--repo-root", repository },
        output,
        error);

    EqualExit(ExitCodes.Success, exitCode, error);
    var text = output.ToString();
    foreach (var argument in new[] { "--base", "$BASE", "--local", "$LOCAL", "--remote", "$REMOTE", "--output", "$MERGED" })
        True(text.Contains(argument, StringComparison.Ordinal));
    True(text.Contains("trustExitCode = true", StringComparison.Ordinal));
    True(text.Contains("config --local", StringComparison.Ordinal));
    True(text.Contains("不会修改 Git 配置", StringComparison.Ordinal));
    True(!File.Exists(gitConfigPath));
}

static void GitLfsMergetoolEndToEnd(string testRoot)
{
    var fixture = CreateGitLfsConflict(Path.Combine(testRoot, "git-lfs-mergetool"));
    var repository = fixture.RepositoryRoot;
    var dataRoot = fixture.DataRoot;
    var tablesPath = fixture.TablesPath;
    var workbookPath = fixture.WorkbookPath;

    var testAssemblyDirectory = Path.GetDirectoryName(typeof(Program).Assembly.Location)!;
    var cliAssemblyPath = typeof(CliApplication).Assembly.Location;
    var runtimeConfigPath = Path.Combine(testAssemblyDirectory, "LubanExcelMerge.Cli.Tests.runtimeconfig.json");
    var diagnosticPath = Path.Combine(repository, "mergetool.jsonl");
    var toolCommand = string.Join(" ", new[]
    {
        "dotnet", "exec", "--runtimeconfig", ShellQuote(runtimeConfigPath), ShellQuote(cliAssemblyPath), "merge",
        "--base", "\"$BASE\"", "--local", "\"$LOCAL\"", "--remote", "\"$REMOTE\"",
        "--output", "\"$MERGED\"", "--repo-root", ShellQuote(repository),
        "--data-root", ShellQuote(dataRoot), "--tables", ShellQuote(tablesPath),
        "--recalculate-with-excel", "never", "--headless", "--log", ShellQuote(diagnosticPath)
    });
    RunProcess(repository, "git", "config", "merge.tool", "luban-test");
    RunProcess(repository, "git", "config", "mergetool.luban-test.cmd", toolCommand);
    RunProcess(repository, "git", "config", "mergetool.luban-test.trustExitCode", "true");

    RunProcess(repository, "git", "mergetool", "--tool=luban-test", "--no-prompt", "--", "ConfigLuban/Datas/Test.xlsx");

    Equal(string.Empty, RunProcess(repository, "git", "diff", "--name-only", "--diff-filter=U").Output.Trim());
    var mergedRows = ReadDataRows(workbookPath);
    Equal("local-a", mergedRows["1"]["A"]);
    Equal("remote-b", mergedRows["1"]["B"]);
    var started = File.ReadLines(diagnosticPath)
        .Select(line => JsonSerializer.Deserialize<JsonElement>(line))
        .First(entry => entry.GetProperty("event").GetString() == "started");
    var files = started.GetProperty("details").GetProperty("files");
    True(new[] { "base", "local", "remote", "merged" }
        .All(name => files.GetProperty(name).GetProperty("exists").GetBoolean()));

    RunProcess(repository, "git", "commit", "-m", "merged by LubanExcelMerge");
    var mergeCommit = RunProcess(repository, "git", "rev-parse", "HEAD").Output.Trim();
    RunProcess(repository, "git", "checkout", "HEAD^1");
    RunProcess(repository, "git", "checkout", mergeCommit);
    var checkedOutRows = ReadDataRows(workbookPath);
    Equal("local-a", checkedOutRows["1"]["A"]);
    Equal("remote-b", checkedOutRows["1"]["B"]);
}

static void GitStagerResolvesAndStagesOnlyMerged(string testRoot)
{
    var fixture = CreateGitLfsConflict(Path.Combine(testRoot, "git-direct-stager"));
    File.AppendAllText(fixture.TablesPath, "# unrelated working-tree change\n", new UTF8Encoding(false));
    File.Delete(fixture.WorkbookPath);
    TestWorkbookFactory.Create(
        fixture.WorkbookPath,
        new[] { new[] { "1", "local-a", "remote-b" } });

    var result = new GitMergedFileStager().Stage(fixture.RepositoryRoot, fixture.WorkbookPath);

    True(result.WasUnmerged);
    Equal("ConfigLuban/Datas/Test.xlsx", result.RelativePath);
    Equal(string.Empty, RunProcess(
        fixture.RepositoryRoot, "git", "ls-files", "--unmerged", "--", result.RelativePath).Output.Trim());
    var cachedPaths = RunProcess(
        fixture.RepositoryRoot, "git", "diff", "--cached", "--name-only").Output
        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    True(cachedPaths.SequenceEqual(new[] { result.RelativePath }, StringComparer.Ordinal));
    True(RunProcess(fixture.RepositoryRoot, "git", "diff", "--name-only").Output
        .Contains("__tables__.csv", StringComparison.Ordinal));
    var indexedWorkbook = RunProcess(
        fixture.RepositoryRoot, "git", "show", $":{result.RelativePath}").Output;
    True(indexedWorkbook.StartsWith("version https://git-lfs.github.com/spec/v1", StringComparison.Ordinal));
    var rows = ReadDataRows(fixture.WorkbookPath);
    Equal("local-a", rows["1"]["A"]);
    Equal("remote-b", rows["1"]["B"]);
}

static void GitStagerDiscoversNestedConfigRepository(string testRoot)
{
    var projectRoot = Path.Combine(testRoot, "nested-project-root");
    Directory.CreateDirectory(projectRoot);
    RunProcess(projectRoot, "git", "init", "-b", "main");
    RunProcess(projectRoot, "git", "config", "user.name", "Luban Excel Merge Test");
    RunProcess(projectRoot, "git", "config", "user.email", "luban-test@example.invalid");
    File.WriteAllText(Path.Combine(projectRoot, ".gitignore"), "/ConfigLuban\n", new UTF8Encoding(false));
    RunProcess(projectRoot, "git", "add", ".gitignore");
    RunProcess(projectRoot, "git", "commit", "-m", "ignore nested config repository");

    var nestedRoot = Path.Combine(projectRoot, "ConfigLuban");
    RunProcess(projectRoot, "git", "init", "-b", "main", nestedRoot);
    True(Directory.Exists(Path.Combine(nestedRoot, ".git")));
    Equal(
        Path.GetFullPath(nestedRoot),
        Path.GetFullPath(RunProcess(nestedRoot, "git", "rev-parse", "--show-toplevel").Output.Trim()));
    var fixture = CreateGitLfsConflict(nestedRoot, "Datas");
    File.Delete(fixture.WorkbookPath);
    TestWorkbookFactory.Create(
        fixture.WorkbookPath,
        new[] { new[] { "1", "local-a", "remote-b" } });

    var result = new GitMergedFileStager().Stage(projectRoot, fixture.WorkbookPath);

    Equal(Path.GetFullPath(fixture.RepositoryRoot), result.RepositoryRoot);
    Equal("Datas/Test.xlsx", result.RelativePath);
    True(result.WasUnmerged);
    Equal(string.Empty, RunProcess(
        fixture.RepositoryRoot, "git", "ls-files", "--unmerged", "--", result.RelativePath).Output.Trim());
    True(RunProcess(fixture.RepositoryRoot, "git", "diff", "--cached", "--name-only").Output
        .Contains(result.RelativePath, StringComparison.Ordinal));
    Equal(string.Empty, RunProcess(projectRoot, "git", "status", "--short").Output.Trim());
}

static GitConflictFixture CreateGitLfsConflict(
    string repository,
    string dataRootRelativePath = "ConfigLuban/Datas")
{
    Directory.CreateDirectory(repository);
    RunProcess(repository, "git", "init", "-b", "main");
    RunProcess(repository, "git", "config", "user.name", "Luban Excel Merge Test");
    RunProcess(repository, "git", "config", "user.email", "luban-test@example.invalid");
    RunProcess(repository, "git", "lfs", "install", "--local");

    var dataRoot = Path.GetFullPath(
        dataRootRelativePath.Replace('/', Path.DirectorySeparatorChar),
        repository);
    Directory.CreateDirectory(dataRoot);
    var tablesPath = Path.Combine(dataRoot, "__tables__.csv");
    var workbookPath = Path.Combine(dataRoot, "Test.xlsx");
    File.WriteAllText(
        Path.Combine(repository, ".gitattributes"),
        "*.xlsx filter=lfs diff=lfs merge=lfs -text\n",
        new UTF8Encoding(false));
    File.WriteAllText(
        tablesPath,
        "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
        "##,说明,,,,,,\n" +
        ",TbTest,Test,TRUE,Test.xlsx,Id,map,c\n",
        new UTF8Encoding(false));
    TestWorkbookFactory.Create(workbookPath, new[] { new[] { "1", "base-a", "base-b" } });
    RunProcess(repository, "git", "add", ".");
    RunProcess(repository, "git", "commit", "-m", "base");
    RunProcess(repository, "git", "branch", "remote");

    File.Delete(workbookPath);
    TestWorkbookFactory.Create(workbookPath, new[] { new[] { "1", "local-a", "base-b" } });
    RunProcess(repository, "git", "add", workbookPath);
    RunProcess(repository, "git", "commit", "-m", "local change");

    RunProcess(repository, "git", "checkout", "remote");
    File.Delete(workbookPath);
    TestWorkbookFactory.Create(workbookPath, new[] { new[] { "1", "base-a", "remote-b" } });
    RunProcess(repository, "git", "add", workbookPath);
    RunProcess(repository, "git", "commit", "-m", "remote change");
    RunProcess(repository, "git", "checkout", "main");
    var merge = RunProcessAllowFailure(repository, "git", "merge", "remote");
    True(merge.ExitCode != 0);
    True(RunProcess(repository, "git", "diff", "--name-only", "--diff-filter=U").Output
        .Contains("Test.xlsx", StringComparison.Ordinal));
    return new GitConflictFixture(repository, dataRoot, tablesPath, workbookPath);
}

static string ShellQuote(string path) => "'" + path.Replace("'", "'\"'\"'", StringComparison.Ordinal) + "'";

static ProcessResult RunProcess(string workingDirectory, string fileName, params string[] arguments) =>
    RunProcessCore(workingDirectory, fileName, arguments, allowFailure: false);

static ProcessResult RunProcessAllowFailure(string workingDirectory, string fileName, params string[] arguments) =>
    RunProcessCore(workingDirectory, fileName, arguments, allowFailure: true);

static ProcessResult RunProcessCore(
    string workingDirectory,
    string fileName,
    IReadOnlyList<string> arguments,
    bool allowFailure)
{
    var startInfo = new ProcessStartInfo(fileName)
    {
        WorkingDirectory = workingDirectory,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true
    };
    foreach (var argument in arguments)
        startInfo.ArgumentList.Add(argument);
    using var process = Process.Start(startInfo)
        ?? throw new InvalidOperationException($"Unable to start {fileName}.");
    var output = process.StandardOutput.ReadToEnd();
    var error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    var result = new ProcessResult(process.ExitCode, output, error);
    if (!allowFailure && result.ExitCode != 0)
        throw new InvalidOperationException(
            $"{fileName} {string.Join(' ', arguments)} failed with {result.ExitCode}.\n{output}\n{error}");
    return result;
}

static void LfsPointerResolves(string testRoot)
{
    var repository = CreateRepository(Path.Combine(testRoot, "lfs-success"));
    var content = Encoding.UTF8.GetBytes("verified lfs object");
    var hash = Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant();
    var objectPath = Path.Combine(repository, ".git", "lfs", "objects", hash[..2], hash.Substring(2, 2), hash);
    Directory.CreateDirectory(Path.GetDirectoryName(objectPath)!);
    File.WriteAllBytes(objectPath, content);
    var pointerPath = Path.Combine(repository, "pointer.xlsx");
    File.WriteAllText(
        pointerPath,
        $"version https://git-lfs.github.com/spec/v1\noid sha256:{hash}\nsize {content.Length}\n",
        new UTF8Encoding(false));

    var resolved = new GitLfsInputResolver().Resolve(pointerPath, repository);
    True(resolved.IsLfsPointer);
    Equal(objectPath, resolved.ContentPath);
}

static void MissingLfsObject(string testRoot)
{
    var repository = CreateRepository(Path.Combine(testRoot, "lfs-missing"));
    var pointerPath = Path.Combine(repository, "pointer.xlsx");
    File.WriteAllText(
        pointerPath,
        "version https://git-lfs.github.com/spec/v1\n" +
        "oid sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\n" +
        "size 10\n",
        new UTF8Encoding(false));
    try
    {
        new GitLfsInputResolver().Resolve(pointerPath, repository);
        throw new InvalidOperationException("Expected LfsObjectNotFoundException.");
    }
    catch (LfsObjectNotFoundException exception)
    {
        True(exception.Message.Contains("git lfs fetch", StringComparison.Ordinal));
    }
}

static void SmallBinaryIsNotLfs(string testRoot)
{
    var path = Path.Combine(testRoot, "small-binary.xlsx");
    File.WriteAllBytes(path, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0xBC, 0xFF, 0x00 });
    True(GitLfsPointer.TryRead(path) is null);
}

static void CliMapsMissingLfsObject(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "cli-lfs-missing"));
    File.WriteAllText(
        scenario.BasePath,
        "version https://git-lfs.github.com/spec/v1\n" +
        "oid sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb\n" +
        "size 100\n",
        new UTF8Encoding(false));
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "a", "b" } });

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.InvalidInput, exitCode, error);
    True(error.ToString().Contains("git lfs fetch", StringComparison.Ordinal));
    True(!File.Exists(scenario.OutputPath));
}

static void NonConflictingMerge(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "merge-success"));
    TestWorkbookFactory.Create(scenario.BasePath, new[]
    {
        new[] { "1", "old-a", "old-b" },
        new[] { "2", "delete-me", "same" }
    });
    TestWorkbookFactory.Create(scenario.LocalPath, new[]
    {
        new[] { "1", "local-a", "old-b" },
        new[] { "2", "delete-me", "same" },
        new[] { "3", "local-new", "local" }
    });
    TestWorkbookFactory.Create(scenario.RemotePath, new[]
    {
        new[] { "1", "old-a", "remote-b" },
        new[] { "4", "remote-new", "remote" }
    });
    var sourceHashes = new[] { HashFile(scenario.BasePath), HashFile(scenario.LocalPath), HashFile(scenario.RemotePath) };

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.Success, exitCode, error);
    True(File.Exists(scenario.OutputPath));
    var rows = ReadDataRows(scenario.OutputPath);
    Equal("local-a", rows["1"]["A"]);
    Equal("remote-b", rows["1"]["B"]);
    True(!rows.ContainsKey("2"));
    Equal("local-new", rows["3"]["A"]);
    Equal("remote-new", rows["4"]["A"]);
    Equal(sourceHashes[0], HashFile(scenario.BasePath));
    Equal(sourceHashes[1], HashFile(scenario.LocalPath));
    Equal(sourceHashes[2], HashFile(scenario.RemotePath));
    True(output.ToString().Contains("写入单元格=1", StringComparison.Ordinal));
    True(output.ToString().Contains("新增记录=1", StringComparison.Ordinal));
    True(output.ToString().Contains("删除记录=1", StringComparison.Ordinal));
}

static void AutomaticMergeResultsIncludeBothSides(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "automatic-edit-location"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "base-a", "base-b" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local-a", "base-b" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "base-a", "remote-b" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.AutomaticEditCount);
    Equal(2, session.AutomaticMergeCount);
    Equal(2, session.ProcessedMergeCount);
    True(session.CanSave);
    var mergeLocations = session.Sheets.Single().AutomaticMergeLocations;
    Equal(2, mergeLocations.Count);
    Equal("C4", mergeLocations[0].DisplayLocation);
    Equal("D4", mergeLocations[1].DisplayLocation);
    var location = session.Sheets.Single().FirstAutomaticEditLocation;
    True(location is not null);
    Equal(3, location!.RowIndex);
    Equal(3, location.ColumnIndex);
    Equal("D4", location.DisplayLocation);

    var deleteScenario = CreateScenario(Path.Combine(testRoot, "automatic-delete-location"));
    TestWorkbookFactory.Create(deleteScenario.BasePath, new[] { new[] { "1", "delete-me", "same" } });
    TestWorkbookFactory.Create(deleteScenario.LocalPath, new[] { new[] { "1", "delete-me", "same" } });
    TestWorkbookFactory.Create(deleteScenario.RemotePath, Array.Empty<string[]>());
    var deleteSession = new LubanMergeCoordinator().Prepare(
        CommandLineParser.Parse(CreateArguments(deleteScenario)));
    var deleteLocation = deleteSession.Sheets.Single().FirstAutomaticEditLocation;
    True(deleteLocation is not null);
    Equal(3, deleteLocation!.RowIndex);
    Equal(1, deleteLocation.ColumnIndex);
    Equal(3, deleteSession.AutomaticMergeCount);
    Equal(0, deleteSession.ProcessedMergeCount);
    True(deleteSession.Sheets.Single().AutomaticMergeLocations
        .Select(item => item.DisplayLocation)
        .SequenceEqual(new[] { "B4", "C4", "D4" }, StringComparer.Ordinal));

    var appendScenario = CreateScenario(Path.Combine(testRoot, "automatic-append-location"));
    TestWorkbookFactory.Create(appendScenario.BasePath, Array.Empty<string[]>());
    TestWorkbookFactory.Create(appendScenario.LocalPath, Array.Empty<string[]>());
    TestWorkbookFactory.Create(appendScenario.RemotePath, new[] { new[] { "2", "new-a", "new-b" } });
    var appendSession = new LubanMergeCoordinator().Prepare(
        CommandLineParser.Parse(CreateArguments(appendScenario)));
    var appendLocation = appendSession.Sheets.Single().FirstAutomaticEditLocation;
    True(appendLocation is not null);
    Equal(3, appendLocation!.RowIndex);
    Equal(1, appendLocation.ColumnIndex);
    Equal(3, appendSession.AutomaticMergeCount);
    Equal(0, appendSession.ProcessedMergeCount);

    var localDeleteScenario = CreateScenario(Path.Combine(testRoot, "automatic-local-delete-location"));
    TestWorkbookFactory.Create(localDeleteScenario.BasePath, new[] { new[] { "3", "delete-me", "same" } });
    TestWorkbookFactory.Create(localDeleteScenario.LocalPath, Array.Empty<string[]>());
    TestWorkbookFactory.Create(localDeleteScenario.RemotePath, new[] { new[] { "3", "delete-me", "same" } });
    var localDeleteSession = new LubanMergeCoordinator().Prepare(
        CommandLineParser.Parse(CreateArguments(localDeleteScenario)));
    Equal(0, localDeleteSession.AutomaticEditCount);
    Equal(3, localDeleteSession.AutomaticMergeCount);
    Equal(0, localDeleteSession.ProcessedMergeCount);
}

static void ModeOneMergesSingleRecord(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "mode-one-automatic"), mode: "one");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "base", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "local", "remote" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(0, session.RemainingConflicts);
    Equal(1, session.ChangedCells);
    session.Save();

    var merged = new OpenXmlWorkbookReader().Read(scenario.OutputPath).GetSheet("Data");
    Equal("local", merged.GetCell("C4")!.Payload.RawValue);
    Equal("remote", merged.GetCell("D4")!.Payload.RawValue);
}

static void ModeOneFieldConflictRequiresResolution(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "mode-one-conflict"), mode: "one");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "base", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.RemainingConflicts);
    Equal(Core.MergeConflictKind.CellChangedDifferently, session.Conflicts.Single().Conflict.Kind);
    session.Conflicts.Single().Resolve(MergeChoice.Remote);
    session.Save();

    Equal("remote", new OpenXmlWorkbookReader().Read(scenario.OutputPath)
        .GetSheet("Data").GetCell("C4")!.Payload.RawValue);
}

static void ModeOneRejectsMultipleRecords(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "mode-one-multiple"), mode: "one");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "base", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[]
    {
        new[] { "1", "local", "same" },
        new[] { "2", "unexpected", "row" }
    });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });

    Throws<UnsafeWorkbookException>(() => new LubanMergeCoordinator().Prepare(
        CommandLineParser.Parse(CreateArguments(scenario))));
}

static void BlankLayoutDoesNotDisplaceAdditions(string testRoot)
{
    var rowScenario = CreateScenario(Path.Combine(testRoot, "blank-layout-row"));
    TestWorkbookFactory.Create(rowScenario.BasePath, new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.Create(rowScenario.LocalPath, new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.Create(rowScenario.RemotePath, new[]
    {
        new[] { "1", "a", "b" },
        new[] { "2", "new-a", "new-b" }
    });
    AddSerializedBlankCell(rowScenario.LocalPath, "Data", "B1000", "1");

    var rowSession = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(rowScenario)));
    rowSession.Save();
    Equal("2", ReadCellValue(rowScenario.OutputPath, "Data", "B5"));
    Equal("new-a", ReadCellValue(rowScenario.OutputPath, "Data", "C5"));
    True(new OpenXmlWorkbookReader().Read(rowScenario.OutputPath).GetSheet("Data").GetCell("B1000") is null);
    True(ReadCellValue(rowScenario.OutputPath, "Data", "B1001") is null);

    var columnScenario = CreateScenario(Path.Combine(testRoot, "blank-layout-column"));
    TestWorkbookFactory.CreateWithSchema(
        columnScenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        columnScenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        columnScenario.RemotePath,
        new[] { "Id", "A", "B", "RemoteField" },
        new[] { "string", "string", "string", "string" },
        new[] { new[] { "1", "a", "b", "remote" } });
    AddSerializedBlankCell(columnScenario.LocalPath, "Data", "E4", "1");
    AddSerializedBlankCell(columnScenario.LocalPath, "Data", "Z1000", "1");

    var columnSession = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(columnScenario)));
    Equal(5, columnSession.Comparison.ColumnHeaders.Count);
    columnSession.Save();
    Equal(4, ParseWorkbookSchema(columnScenario.OutputPath).FindField("RemoteField")!.ColumnIndex);
    Equal("remote", ReadCellValue(columnScenario.OutputPath, "Data", "E4"));
    var columnOutput = new OpenXmlWorkbookReader().Read(columnScenario.OutputPath).GetSheet("Data");
    Equal<string?>(null, columnOutput.GetCell("E4")!.StyleIndex);
    True(columnOutput.GetCell("Z1000") is null);
    True(ReadCellValue(columnScenario.OutputPath, "Data", "AA4") is null);
}

static void LargeTablePerformanceSmoke(string testRoot)
{
    const int recordCount = 10_000;
    var scenario = CreateScenario(Path.Combine(testRoot, "large-table-performance"));
    var baseRows = Enumerable.Range(1, recordCount)
        .Select(index => new[] { index.ToString(), $"a-{index}", $"b-{index}" })
        .ToArray();
    var remoteRows = Enumerable.Range(1, recordCount)
        .Select(index => new[]
        {
            index.ToString(),
            $"a-{index}",
            index % 250 == 0 ? $"remote-{index}" : $"b-{index}"
        })
        .ToArray();
    TestWorkbookFactory.Create(scenario.BasePath, baseRows);
    TestWorkbookFactory.Create(scenario.LocalPath, baseRows);
    TestWorkbookFactory.Create(scenario.RemotePath, remoteRows);

    var stopwatch = Stopwatch.StartNew();
    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    var observedPreparationMilliseconds = stopwatch.ElapsedMilliseconds;

    stopwatch.Restart();
    var tables = Enum.GetValues<MergeGridSide>()
        .Select(side => session.Comparison.CreateTable(side))
        .ToArray();
    var tableCreationMilliseconds = stopwatch.ElapsedMilliseconds;

    stopwatch.Restart();
    var locations = session.Sheets.Single().AutomaticMergeLocations;
    var locationIndexMilliseconds = stopwatch.ElapsedMilliseconds;

    Equal(40, session.AutomaticEditCount);
    Equal(40, session.AutomaticMergeCount);
    True(tables.All(table => table.Rows.Count == recordCount + 3));
    Equal(40, locations.Count);
    True(observedPreparationMilliseconds < 30_000);
    True(tableCreationMilliseconds < 2_000);
    True(locationIndexMilliseconds < 2_000);
    Console.WriteLine(
        $"BENCH rows={recordCount} prepare={observedPreparationMilliseconds}ms " +
        $"reported={session.PreparationTimings.TotalMilliseconds}ms " +
        $"tables={tableCreationMilliseconds}ms locations={locationIndexMilliseconds}ms");
}

static void MultiSheetAutomaticMerge(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "multi-sheet-automatic"));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.BasePath,
        ("Alpha", new[] { new[] { "1", "base-a", "base-b" } }),
        ("Beta", new[] { new[] { "10", "base-x", "base-y" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.LocalPath,
        ("Alpha", new[] { new[] { "1", "local-a", "base-b" } }),
        ("Beta", new[] { new[] { "10", "base-x", "base-y" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.RemotePath,
        ("Alpha", new[] { new[] { "1", "base-a", "remote-b" } }),
        ("Beta", new[] { new[] { "10", "base-x", "remote-y" } }));
    var prepared = new LubanMergeCoordinator().Prepare(
        CommandLineParser.Parse(CreateArguments(scenario)));
    var automaticLocations = prepared.Sheets
        .SelectMany(sheet => sheet.AutomaticMergeLocations.Select(location =>
            (sheet.SheetName, Location: location)))
        .ToArray();
    Equal(3, automaticLocations.Length);
    Equal("Alpha", automaticLocations[0].SheetName);
    Equal("C4", automaticLocations[0].Location.DisplayLocation);
    Equal("Alpha", automaticLocations[1].SheetName);
    Equal("D4", automaticLocations[1].Location.DisplayLocation);
    Equal("Beta", automaticLocations[2].SheetName);
    Equal("D4", automaticLocations[2].Location.DisplayLocation);
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.Success, exitCode, error);
    var workbook = new OpenXmlWorkbookReader().Read(scenario.OutputPath);
    True(workbook.Sheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "Alpha", "Beta" }, StringComparer.Ordinal));
    var alpha = ReadDataRows(scenario.OutputPath, "Alpha");
    var beta = ReadDataRows(scenario.OutputPath, "Beta");
    Equal("local-a", alpha["1"]["A"]);
    Equal("remote-b", alpha["1"]["B"]);
    Equal("remote-y", beta["10"]["B"]);
    True(output.ToString().Contains("Alpha, Beta", StringComparison.Ordinal));
}

static void MultiSheetRequiresEveryConflictResolved(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "multi-sheet-conflicts"));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.BasePath,
        ("Alpha", new[] { new[] { "1", "base-a", "same" } }),
        ("Beta", new[] { new[] { "2", "base-b", "same" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.LocalPath,
        ("Alpha", new[] { new[] { "1", "local-a", "same" } }),
        ("Beta", new[] { new[] { "2", "local-b", "same" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.RemotePath,
        ("Alpha", new[] { new[] { "1", "remote-a", "same" } }),
        ("Beta", new[] { new[] { "2", "remote-b", "same" } }));

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(2, session.Sheets.Count);
    Equal(2, session.RemainingConflicts);
    Equal(1, session.Sheets[0].RemainingConflicts);
    Equal(1, session.Sheets[1].RemainingConflicts);
    session.Sheets[0].Conflicts[0].Resolve(MergeChoice.Local);
    True(!session.CanSave);
    try
    {
        session.Save();
        throw new InvalidOperationException("Expected unresolved multi-sheet save to fail.");
    }
    catch (InvalidOperationException exception)
    {
        True(exception.Message.Contains("1 个冲突未解决", StringComparison.Ordinal));
    }
    True(!File.Exists(scenario.OutputPath));

    session.Sheets[1].Conflicts[0].Resolve(MergeChoice.Remote);
    True(session.CanSave);
    session.Save();
    Equal("local-a", ReadDataRows(scenario.OutputPath, "Alpha")["1"]["A"]);
    Equal("remote-b", ReadDataRows(scenario.OutputPath, "Beta")["2"]["A"]);
}

static void MultiSheetNamesMustMatch(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "multi-sheet-name-mismatch"));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.BasePath,
        ("Alpha", new[] { new[] { "1", "a", "b" } }),
        ("Beta", new[] { new[] { "2", "a", "b" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.LocalPath,
        ("Alpha", new[] { new[] { "1", "a", "b" } }),
        ("Beta", new[] { new[] { "2", "a", "b" } }));
    TestWorkbookFactory.CreateMultiSheet(
        scenario.RemotePath,
        ("Alpha", new[] { new[] { "1", "a", "b" } }),
        ("Gamma", new[] { new[] { "2", "a", "b" } }));
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.UnsafeWorkbook, exitCode, error);
    True(error.ToString().Contains("名称、数量或顺序不一致", StringComparison.Ordinal));
    True(!File.Exists(scenario.OutputPath));
}

static void MetadataChangeRequiresReview(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "metadata-review"));
    var records = new[] { new[] { "1", "same-a", "same-b" } };
    TestWorkbookFactory.Create(scenario.BasePath, records);
    TestWorkbookFactory.Create(scenario.LocalPath, records);
    TestWorkbookFactory.Create(scenario.RemotePath, records);
    SetWorkbookCell(scenario.RemotePath, "Data", "D2", "远端字段说明");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.MetadataChangeCount);
    Equal(1, session.RemainingConflicts);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.MetadataChanged, conflict.Conflict.Kind);
    Equal("D2", conflict.Conflict.FieldName);
    Equal(2, conflict.RowNumber);
    Equal(1, conflict.GridRowIndex);
    Equal(3, conflict.GridColumnIndex);
    Equal(MergeGridCellState.Metadata,
        session.Comparison.CreateTable(MergeGridSide.Remote).Rows[1].Cells[3].State);
    True(!session.CanSave);

    conflict.Resolve(MergeChoice.Remote);
    Equal("远端字段说明",
        session.Comparison.CreateTable(MergeGridSide.Merged).Rows[1].Cells[3].DisplayValue);
    Equal(MergeGridCellState.Modified,
        session.Comparison.CreateTable(MergeGridSide.Merged).Rows[1].Cells[3].State);
    session.Save();
    Equal("远端字段说明", ReadCellValue(scenario.OutputPath, "Data", "D2"));
}

static void MultiSheetMetadataChangesGateSave(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "multi-sheet-metadata-review"));
    var sheets = new[]
    {
        ("Alpha", (IReadOnlyList<string[]>)new[] { new[] { "1", "a", "b" } }),
        ("Beta", (IReadOnlyList<string[]>)new[] { new[] { "2", "a", "b" } })
    };
    TestWorkbookFactory.CreateMultiSheet(scenario.BasePath, sheets);
    TestWorkbookFactory.CreateMultiSheet(scenario.LocalPath, sheets);
    TestWorkbookFactory.CreateMultiSheet(scenario.RemotePath, sheets);
    SetWorkbookCell(scenario.RemotePath, "Alpha", "D2", "Alpha 远端说明");
    SetWorkbookCell(scenario.RemotePath, "Beta", "D2", "Beta 远端说明");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(2, session.MetadataChangeCount);
    Equal(1, session.Sheets[0].MetadataChangeCount);
    Equal(1, session.Sheets[1].MetadataChangeCount);
    session.Sheets[0].Conflicts.Single().Resolve(MergeChoice.Remote);
    True(!session.CanSave);
    session.Sheets[1].Conflicts.Single().Resolve(MergeChoice.Remote);
    True(session.CanSave);
    session.Save();
    Equal("Alpha 远端说明", ReadCellValue(scenario.OutputPath, "Alpha", "D2"));
    Equal("Beta 远端说明", ReadCellValue(scenario.OutputPath, "Beta", "D2"));
}

static void MultiSheetRemoteAppendedFieldsSaveEverySheet(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "multi-sheet-appended-fields"));
    var sheets = new[]
    {
        ("Alpha", (IReadOnlyList<string[]>)new[] { new[] { "1", "a", "b" } }),
        ("Beta", (IReadOnlyList<string[]>)new[] { new[] { "2", "a", "b" } })
    };
    TestWorkbookFactory.CreateMultiSheet(scenario.BasePath, sheets);
    TestWorkbookFactory.CreateMultiSheet(scenario.LocalPath, sheets);
    TestWorkbookFactory.CreateMultiSheet(scenario.RemotePath, sheets);
    foreach (var (sheetName, value) in new[] { ("Alpha", "alpha-extra"), ("Beta", "beta-extra") })
    {
        SetWorkbookCell(scenario.RemotePath, sheetName, "E1", "RemoteField");
        SetWorkbookCell(scenario.RemotePath, sheetName, "E2", "远端新增字段");
        SetWorkbookCell(scenario.RemotePath, sheetName, "E3", "string");
        SetWorkbookCell(scenario.RemotePath, sheetName, "E4", value);
    }

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    Equal(2, session.Sheets.Count);
    session.Save();
    Equal("alpha-extra", ReadCellValue(scenario.OutputPath, "Alpha", "E4"));
    Equal("beta-extra", ReadCellValue(scenario.OutputPath, "Beta", "E4"));
}

static void RemoteOnlyAppendedFieldMerges(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "remote-appended-field"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base-a", "base-b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "local-a", "base-b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B", "RemoteField" },
        new[] { "string", "string", "string", "int" },
        new[]
        {
            new[] { "1", "base-a", "remote-b", "42" },
            new[] { "2", "new-a", "new-b", "84" }
        });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.CanSave);
    True(session.AutomaticEditCount >= 3);
    var remoteGrid = session.Comparison.CreateTable(MergeGridSide.Remote);
    var mergedGrid = session.Comparison.CreateTable(MergeGridSide.Merged);
    Equal("RemoteField", remoteGrid.Rows[0].Cells[4].DisplayValue);
    Equal("RemoteField", mergedGrid.Rows[0].Cells[4].DisplayValue);
    Equal(MergeGridCellState.Added, mergedGrid.Rows[0].Cells[4].State);

    session.Save();
    var outputSchema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("RemoteField", outputSchema.Fields.Single(field => field.Name == "RemoteField").Name);
    Equal("int", outputSchema.FindField("RemoteField")!.TypeName);
    Equal("local-a", ReadCellValue(scenario.OutputPath, "Data", "C4"));
    Equal("remote-b", ReadCellValue(scenario.OutputPath, "Data", "D4"));
    Equal("42", ReadCellValue(scenario.OutputPath, "Data", "E4"));
    Equal("2", ReadCellValue(scenario.OutputPath, "Data", "B5"));
    Equal("84", ReadCellValue(scenario.OutputPath, "Data", "E5"));
}

static void DifferentAppendedFieldsFormUnion(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "different-appended-fields"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "Value" },
        new[] { "string", "string" },
        new[] { new[] { "1", "base" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "Value", "LocalField" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base", "local-value" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "Value", "RemoteField" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base", "remote-value" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    session.Save();
    var outputSchema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("LocalField", outputSchema.Fields[2].Name);
    Equal("RemoteField", outputSchema.Fields[3].Name);
    Equal("local-value", ReadCellValue(scenario.OutputPath, "Data", "D4"));
    Equal("remote-value", ReadCellValue(scenario.OutputPath, "Data", "E4"));
}

static void SameAppendedFieldDataConflictIsResolvable(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "same-appended-field-conflict"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "Value" },
        new[] { "string", "string" },
        new[] { new[] { "1", "base" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "Value", "SharedField" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base", "local-value" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "Value", "SharedField" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base", "remote-value" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.RemainingConflicts);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.CellChangedDifferently, conflict.Conflict.Kind);
    Equal("SharedField", conflict.Conflict.FieldName);
    conflict.Resolve(MergeChoice.Remote);
    session.Save();
    Equal("remote-value", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void SameAppendedFieldTypeConflictRequiresReview(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "same-appended-field-type"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "Value" },
        new[] { "string", "string" },
        new[] { new[] { "1", "base" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "Value", "SharedField" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "base", "1" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "Value", "SharedField" },
        new[] { "string", "string", "int" },
        new[] { new[] { "1", "base", "1" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.RemainingConflicts);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.MetadataChanged, conflict.Conflict.Kind);
    Equal("D1", conflict.Conflict.FieldName);
    conflict.Resolve(MergeChoice.Remote);
    session.Save();
    Equal("int", ParseWorkbookSchema(scenario.OutputPath).FindField("SharedField")!.TypeName);
}

static void RemoteInsertedFieldReordersExistingFields(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "inserted-field-inside-layout"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "local-a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "Inserted", "A", "B" },
        new[] { "string", "string", "string", "string" },
        new[] { new[] { "1", "new", "a", "remote-b" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("Id", schema.Fields[0].Name);
    Equal("Inserted", schema.Fields[1].Name);
    Equal("A", schema.Fields[2].Name);
    Equal("B", schema.Fields[3].Name);
    Equal("new", ReadCellValue(scenario.OutputPath, "Data", "C4"));
    Equal("local-a", ReadCellValue(scenario.OutputPath, "Data", "D4"));
    Equal("remote-b", ReadCellValue(scenario.OutputPath, "Data", "E4"));
}

static void RemoteInsertedFieldMovesFormulaField(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "inserted-field-moves-formula"));
    foreach (var path in new[] { scenario.BasePath, scenario.LocalPath })
    {
        TestWorkbookFactory.CreateWithSchema(
            path,
            new[] { "Id", "A", "B" },
            new[] { "string", "string", "int" },
            new[] { new[] { "1", "a", "" } });
        SetWorkbookFormula(path, "Data", "D4", "LEN(C4)", "1");
    }
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "Inserted", "A", "B" },
        new[] { "string", "string", "string", "int" },
        new[] { new[] { "1", "new", "a", "" } });
    SetWorkbookFormula(scenario.RemotePath, "Data", "E4", "LEN(D4)", "1");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    var result = session.Save();
    Equal(WorkbookRecalculationStatus.SourceCachePreservedUnverified, result.RecalculationStatus);
    var output = new OpenXmlWorkbookReader().Read(scenario.OutputPath).GetSheet("Data");
    Equal("new", output.GetCell("C4")!.Payload.RawValue);
    Equal("a", output.GetCell("D4")!.Payload.RawValue);
    Equal("LEN(D4)", output.GetCell("E4")!.Payload.FormulaText);
    Equal("1", output.GetCell("E4")!.Payload.CachedValue);
}

static void LocalInsertedFieldDefinesMergedLayout(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "local-inserted-field"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "base-b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "Inserted", "A", "B" },
        new[] { "string", "string", "string", "string" },
        new[] { new[] { "1", "local-new", "a", "base-b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "remote-b" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("Inserted", schema.Fields[1].Name);
    Equal("A", schema.Fields[2].Name);
    Equal("B", schema.Fields[3].Name);
    Equal("local-new", ReadCellValue(scenario.OutputPath, "Data", "C4"));
    Equal("remote-b", ReadCellValue(scenario.OutputPath, "Data", "E4"));
}

static void RemoteReorderedExistingFieldsMerge(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "reordered-existing-fields"));
    var records = new[] { new[] { "1", "a", "b" } };
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "B", "A" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "b", "a" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("Id", schema.Fields[0].Name);
    Equal("B", schema.Fields[1].Name);
    Equal("A", schema.Fields[2].Name);
    Equal("b", ReadCellValue(scenario.OutputPath, "Data", "C4"));
    Equal("a", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void RemoteDeletedExistingFieldMerges(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "deleted-existing-field"));
    var records = new[] { new[] { "1", "a", "b" } };
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A" },
        new[] { "string", "string" },
        new[] { new[] { "1", "a" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    Equal(string.Empty, session.Comparison.CreateTable(MergeGridSide.Merged)
        .Rows.Single(row => row.RecordKey == "1").Cells[3].DisplayValue);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    True(schema.FindField("B") is null);
    Equal(2, schema.Fields.Count);
    True(ReadCellValue(scenario.OutputPath, "Data", "D4") is null);
}

static void RemoteExistingFieldTypeChangeMerges(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "existing-field-type-change"));
    var records = new[] { new[] { "1", "a", "1" } };
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        records);
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "int" },
        records);

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    session.Save();
    Equal("int", ParseWorkbookSchema(scenario.OutputPath).FindField("B")!.TypeName);
    Equal("1", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void RemoteExistingFieldRenameMerges(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "metadata-field-definition"));
    var records = new[] { new[] { "1", "a", "b" } };
    TestWorkbookFactory.Create(scenario.BasePath, records);
    TestWorkbookFactory.Create(scenario.LocalPath, records);
    TestWorkbookFactory.Create(scenario.RemotePath, records);
    SetWorkbookCell(scenario.RemotePath, "Data", "D1", "RenamedField");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    True(schema.FindField("B") is null);
    Equal("string", schema.FindField("RenamedField")!.TypeName);
    Equal("b", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void LocalPrimaryKeyFieldRenamePreservesIdentity(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "local-key-field-rename"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Key", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "local-a", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "remote-b" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    True(schema.FindField("Id") is null);
    True(schema.FindField("Key") is not null);
    Equal("local-a", ReadCellValue(scenario.OutputPath, "Data", "C4"));
    Equal("remote-b", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void ColumnDeleteModifyConflictIsResolvable(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "column-delete-modify"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "7" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A" },
        new[] { "string", "string" },
        new[] { new[] { "1", "a" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "int" },
        new[] { new[] { "1", "a", "8" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.DeleteModify, conflict.Conflict.Kind);
    True(conflict.LocalValue.Contains("已删除", StringComparison.Ordinal));
    Equal(string.Empty, session.Comparison.CreateTable(MergeGridSide.Merged)
        .Rows.Single(row => row.RecordKey == "1").Cells[3].DisplayValue);
    conflict.Resolve(MergeChoice.Remote);
    Equal("8", session.Comparison.CreateTable(MergeGridSide.Merged)
        .Rows.Single(row => row.RecordKey == "1").Cells[3].DisplayValue);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    Equal("int", schema.FindField("B")!.TypeName);
    Equal("8", ReadCellValue(scenario.OutputPath, "Data", "D4"));
}

static void ColumnModifyModifyConflictIsResolvable(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "column-modify-modify"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "string" },
        new[] { new[] { "1", "a", "1" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "int" },
        new[] { new[] { "1", "a", "1" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "A", "B" },
        new[] { "string", "string", "bool" },
        new[] { new[] { "1", "a", "1" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.MetadataChanged, conflict.Conflict.Kind);
    conflict.Resolve(MergeChoice.Remote);
    session.Save();
    Equal("bool", ParseWorkbookSchema(scenario.OutputPath).FindField("B")!.TypeName);
}

static void ColumnMoveMoveConflictIsResolvable(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "column-move-move"));
    TestWorkbookFactory.CreateWithSchema(
        scenario.BasePath,
        new[] { "Id", "", "A", "", "B" },
        new[] { "string", "", "string", "", "string" },
        new[] { new[] { "1", "", "a", "", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.LocalPath,
        new[] { "Id", "A", "", "", "B" },
        new[] { "string", "string", "", "", "string" },
        new[] { new[] { "1", "a", "", "", "b" } });
    TestWorkbookFactory.CreateWithSchema(
        scenario.RemotePath,
        new[] { "Id", "", "", "A", "B" },
        new[] { "string", "", "", "string", "string" },
        new[] { new[] { "1", "", "", "a", "b" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(1, session.RemainingConflicts);
    True(session.RequiresStructuralChangeConfirmation);
    var conflict = session.Conflicts.Single();
    Equal(Core.MergeConflictKind.MetadataChanged, conflict.Conflict.Kind);
    True(conflict.LocalValue.Contains("C列", StringComparison.Ordinal));
    True(conflict.RemoteValue.Contains("E列", StringComparison.Ordinal));
    conflict.Resolve(MergeChoice.Remote);
    session.Save();
    var schema = ParseWorkbookSchema(scenario.OutputPath);
    Equal(4, schema.FindField("A")!.ColumnIndex);
    Equal("a", ReadCellValue(scenario.OutputPath, "Data", "E4"));
    True(ReadCellValue(scenario.OutputPath, "Data", "C4") is null);
}

static LubanExcelMerge.Luban.LubanSchema ParseWorkbookSchema(string path)
{
    var sheet = new OpenXmlWorkbookReader().Read(path).Sheets.Single();
    var rows = sheet.Rows.Select(row =>
    {
        var maximumColumn = row.Cells.Select(cell => cell.ColumnIndex).DefaultIfEmpty(-1).Max();
        var values = new string?[maximumColumn + 1];
        foreach (var cell in row.Cells)
            values[cell.ColumnIndex] = cell.Payload.RawValue;
        return new LubanExcelMerge.Luban.LubanRawRow(row.RowNumber, values);
    }).ToArray();
    return LubanExcelMerge.Luban.LubanSchemaParser.Parse(rows);
}

static void CommentedDataRowsAreIgnored(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "commented-data-rows"));
    var records = new[]
    {
        new[] { "1", "disabled", "old" },
        new[] { "1", "active", "same" },
        new[] { "##valid-id", "active", "same" }
    };
    foreach (var path in new[] { scenario.BasePath, scenario.LocalPath, scenario.RemotePath })
    {
        TestWorkbookFactory.Create(path, records);
        SetWorkbookCell(path, "Data", "A4", "##");
    }

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(0, session.RemainingConflicts);
    Equal(0, session.AutomaticEditCount);
    True(session.CanSave);
    True(session.Comparison.CreateTable(MergeGridSide.Local).Rows.Any(row => row.RecordKey == "##valid-id"));
}

static void SetWorkbookCell(string path, string sheetName, string address, string value) =>
    new OpenXmlWorkbookEditor().Apply(
        path,
        new WorkbookEdit[]
        {
            new SetCellEdit(sheetName, address, new Core.CellPayload(Core.CellValueKind.String, value))
        });

static void AddSerializedBlankCell(
    string path,
    string sheetName,
    string address,
    string? styleIndex = null) =>
    new OpenXmlWorkbookEditor().Apply(
        path,
        new WorkbookEdit[]
        {
            new SetCellEdit(sheetName, address, Core.CellPayload.Blank, styleIndex)
        });

static void SetWorkbookFormula(
    string path,
    string sheetName,
    string address,
    string formula,
    string cachedValue) =>
    new OpenXmlWorkbookEditor().Apply(
        path,
        new WorkbookEdit[]
        {
            new SetCellEdit(
                sheetName,
                address,
                new Core.CellPayload(
                    Core.CellValueKind.Formula,
                    FormulaText: formula,
                    CachedValue: cachedValue))
        });

static string? ReadCellValue(string path, string sheetName, string address) =>
    new OpenXmlWorkbookReader().Read(path).GetSheet(sheetName).GetCell(address)?.Payload.RawValue;

static void ConflictPreservesMerged(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "merge-conflict"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "99", "sentinel", "keep" } });
    var before = HashFile(scenario.OutputPath);

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.UnresolvedConflicts, exitCode, error);
    Equal(before, HashFile(scenario.OutputPath));
    True(error.ToString().Contains("未被修改", StringComparison.Ordinal));
}

static void InteractiveCellResolution(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "interactive-cell"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old-a", "old-b" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local-a", "old-b" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote-a", "remote-b" } });

    var options = CommandLineParser.Parse(CreateArguments(scenario));
    var session = new LubanMergeCoordinator().Prepare(options);
    Equal(1, session.RemainingConflicts);
    Equal(1, session.AutomaticEditCount);
    Equal(1, session.AutomaticMergeCount);
    Equal(1, session.ProcessedMergeCount);
    True(!session.CanSave);
    Equal("old-a", session.Conflicts[0].BaseValue);
    Equal("local-a", session.Conflicts[0].LocalValue);
    Equal("remote-a", session.Conflicts[0].RemoteValue);
    Equal(3, session.Conflicts[0].GridRowIndex);
    Equal(2, session.Conflicts[0].GridColumnIndex);
    var baseGrid = session.Comparison.CreateTable(MergeGridSide.Base);
    var remoteGrid = session.Comparison.CreateTable(MergeGridSide.Remote);
    var initialMergedGrid = session.Comparison.CreateTable(MergeGridSide.Merged);
    var baseRow = baseGrid.Rows.Single(row => row.RecordKey == "1");
    var remoteRow = remoteGrid.Rows.Single(row => row.RecordKey == "1");
    var initialMergedRow = initialMergedGrid.Rows.Single(row => row.RecordKey == "1");
    Equal(MergeGridCellState.Conflict, baseRow.Cells[2].State);
    Equal(MergeGridCellState.Modified, remoteRow.Cells[3].State);
    Equal(MergeGridCellState.Normal, baseRow.Cells[3].State);
    Equal(MergeGridCellState.Modified, initialMergedRow.Cells[3].State);
    Equal("local-a", initialMergedRow.Cells[2].DisplayValue);
    Equal("remote-b", initialMergedRow.Cells[3].DisplayValue);
    Equal(MergeGridCellState.Conflict, initialMergedRow.Cells[2].State);

    session.Conflicts[0].Resolve(MergeChoice.Remote);
    True(session.CanSave);
    Equal(2, session.ProcessedMergeCount);
    True(session.Sheets.Single().ProcessedMergeLocations
        .Select(item => item.DisplayLocation)
        .SequenceEqual(new[] { "C4", "D4" }, StringComparer.Ordinal));
    var resolvedBaseRow = session.Comparison.CreateTable(MergeGridSide.Base).Rows.Single(row => row.RecordKey == "1");
    var resolvedLocalRow = session.Comparison.CreateTable(MergeGridSide.Local).Rows.Single(row => row.RecordKey == "1");
    var resolvedRemoteRow = session.Comparison.CreateTable(MergeGridSide.Remote).Rows.Single(row => row.RecordKey == "1");
    var resolvedMergedRow = session.Comparison.CreateTable(MergeGridSide.Merged).Rows.Single(row => row.RecordKey == "1");
    Equal("remote-a", resolvedMergedRow.Cells[2].DisplayValue);
    Equal(MergeGridCellState.Modified, resolvedBaseRow.Cells[2].State);
    Equal(MergeGridCellState.Modified, resolvedLocalRow.Cells[2].State);
    Equal(MergeGridCellState.Modified, resolvedRemoteRow.Cells[2].State);
    Equal(MergeGridCellState.Modified, resolvedMergedRow.Cells[2].State);
    session.Save();
    var rows = ReadDataRows(scenario.OutputPath);
    Equal("remote-a", rows["1"]["A"]);
    Equal("remote-b", rows["1"]["B"]);

    session.Conflicts[0].ClearResolution();
    True(!session.CanSave);
    Equal(1, session.ProcessedMergeCount);
    Equal(MergeGridCellState.Conflict,
        session.Comparison.CreateTable(MergeGridSide.Merged).Rows.Single(row => row.RecordKey == "1").Cells[2].State);
}

static void InsertedMetadataRowRequiresReview(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "metadata-row-insert"));
    var records = new[] { new[] { "1", "same-a", "same-b" } };
    TestWorkbookFactory.Create(scenario.BasePath, records);
    TestWorkbookFactory.Create(scenario.LocalPath, records);
    TestWorkbookFactory.Create(scenario.RemotePath, records);
    new OpenXmlWorkbookEditor().Apply(scenario.RemotePath, new WorkbookEdit[]
    {
        new ReplaceMetadataRowsEdit("Data", 1, 3, new[]
        {
            MetadataRow("##var", "Id", "A", "B"),
            MetadataRow("##", "编号", "新增说明", "字段B"),
            MetadataRow("##", "编号", "字段A", "字段B"),
            MetadataRow("##type", "string", "string", "string")
        })
    });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.MetadataChangeCount);
    Equal(1, session.RemainingConflicts);
    True(session.Comparison.CreateTable(MergeGridSide.Remote).Rows
        .Take(4)
        .SelectMany(row => row.Cells)
        .Any(cell => cell.State == MergeGridCellState.Metadata));
    session.Conflicts.Single().Resolve(MergeChoice.Remote);
    session.Save();
    Equal("新增说明", ReadCellValue(scenario.OutputPath, "Data", "C2"));
    Equal("1", ReadCellValue(scenario.OutputPath, "Data", "B5"));

    static RowWrite MetadataRow(params string[] values) => new(values
        .Select((value, index) => new CellWrite(index, new Core.CellPayload(Core.CellValueKind.String, value)))
        .ToArray());
}

static void DeletedMetadataRowRequiresReview(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "metadata-row-delete"));
    var records = new[] { new[] { "1", "same-a", "same-b" } };
    TestWorkbookFactory.Create(scenario.BasePath, records);
    TestWorkbookFactory.Create(scenario.LocalPath, records);
    TestWorkbookFactory.Create(scenario.RemotePath, records);
    new OpenXmlWorkbookEditor().Apply(scenario.RemotePath, new WorkbookEdit[]
    {
        new ReplaceMetadataRowsEdit("Data", 1, 3, new[]
        {
            MetadataRow("##var", "Id", "A", "B"),
            MetadataRow("##type", "string", "string", "string")
        })
    });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.RemainingConflicts);
    session.Conflicts.Single().Resolve(MergeChoice.Remote);
    session.Save();
    Equal("##type", ReadCellValue(scenario.OutputPath, "Data", "A2"));
    Equal("1", ReadCellValue(scenario.OutputPath, "Data", "B3"));

    static RowWrite MetadataRow(params string[] values) => new(values
        .Select((value, index) => new CellWrite(index, new Core.CellPayload(Core.CellValueKind.String, value)))
        .ToArray());
}

static void InteractiveDeleteModifyResolution(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "interactive-delete-modify"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, Array.Empty<string[]>());
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.RemainingConflicts);
    Equal(Core.MergeConflictKind.DeleteModify, session.Conflicts[0].Conflict.Kind);
    session.Conflicts[0].Resolve(MergeChoice.Remote);
    session.Save();
    var rows = ReadDataRows(scenario.OutputPath);
    Equal("remote", rows["1"]["A"]);
}

static void ComparisonGridAlignsRows(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "comparison-alignment"));
    TestWorkbookFactory.Create(scenario.BasePath, new[]
    {
        new[] { "1", "same", "same" },
        new[] { "2", "delete", "same" }
    });
    TestWorkbookFactory.Create(scenario.LocalPath, new[]
    {
        new[] { "1", "same", "same" },
        new[] { "2", "delete", "same" },
        new[] { "3", "local-add", "same" }
    });
    TestWorkbookFactory.Create(scenario.RemotePath, new[]
    {
        new[] { "1", "same", "same" },
        new[] { "4", "remote-add", "same" }
    });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    var baseGrid = session.Comparison.CreateTable(MergeGridSide.Base);
    var localGrid = session.Comparison.CreateTable(MergeGridSide.Local);
    var remoteGrid = session.Comparison.CreateTable(MergeGridSide.Remote);
    var mergedGrid = session.Comparison.CreateTable(MergeGridSide.Merged);
    Equal(baseGrid.Rows.Count, localGrid.Rows.Count);
    Equal(baseGrid.Rows.Count, remoteGrid.Rows.Count);
    Equal(baseGrid.Rows.Count, mergedGrid.Rows.Count);
    Equal(MergeGridCellState.Added, localGrid.Rows.Single(row => row.RecordKey == "3").Cells[1].State);
    Equal(MergeGridCellState.Normal, baseGrid.Rows.Single(row => row.RecordKey == "3").Cells[1].State);
    Equal(MergeGridCellState.Added, remoteGrid.Rows.Single(row => row.RecordKey == "4").Cells[1].State);
    Equal(MergeGridCellState.Deleted, remoteGrid.Rows.Single(row => row.RecordKey == "2").Cells[1].State);
    Equal(MergeGridCellState.Deleted, mergedGrid.Rows.Single(row => row.RecordKey == "2").Cells[1].State);
}

static void FormulaRecalculationModeReachesSave(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "formula-recalculation"));
    TestWorkbookFactory.CreateWithFormula(scenario.BasePath, "old");
    TestWorkbookFactory.CreateWithFormula(scenario.LocalPath, "old");
    TestWorkbookFactory.CreateWithFormula(scenario.RemotePath, "remote");
    var recalculator = new FakeWorkbookRecalculator();
    var coordinator = new LubanMergeCoordinator(
        saver: new AtomicWorkbookSaver(recalculator: recalculator));
    var session = coordinator.Prepare(CommandLineParser.Parse(CreateArguments(scenario, "auto")));
    var result = session.Save();

    Equal(1, recalculator.CallCount);
    Equal(WorkbookRecalculationStatus.Completed, result.RecalculationStatus);
    Equal("remote", ReadDataRows(scenario.OutputPath)["1"]["A"]);
}

static void ConfigDrivesFormulaRecalculation(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "config-formula-recalculation"));
    TestWorkbookFactory.CreateWithFormula(scenario.BasePath, "old");
    TestWorkbookFactory.CreateWithFormula(scenario.LocalPath, "old");
    TestWorkbookFactory.CreateWithFormula(scenario.RemotePath, "remote");
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        "{\"formulaRecalculation\":\"auto\"}",
        new UTF8Encoding(false));
    var recalculator = new FakeWorkbookRecalculator();
    var coordinator = new LubanMergeCoordinator(
        saver: new AtomicWorkbookSaver(recalculator: recalculator));
    var arguments = CreateArguments(scenario)[..^2];

    coordinator.Prepare(CommandLineParser.Parse(arguments)).Save();

    Equal(1, recalculator.CallCount);
}

static void StrictConfigRejectsUnknownFields(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "strict-config"));
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        "{\"formulaRecalculation\":\"never\",\"typoField\":true}",
        new UTF8Encoding(false));

    try
    {
        LubanMergeConfigurationLoader.Apply(
            CommandLineParser.Parse(CreateArguments(scenario)[..^2]));
        throw new InvalidOperationException("Expected MergeInputException.");
    }
    catch (MergeInputException exception)
    {
        True(exception.Message.Contains("typoField", StringComparison.Ordinal));
    }
}

static void KeyOverrideChangesRecordIdentity(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "key-override"));
    var baseRows = new[]
    {
        new[] { "1", "key-a", "old-a" },
        new[] { "1", "key-b", "old-b" }
    };
    TestWorkbookFactory.Create(scenario.BasePath, baseRows);
    TestWorkbookFactory.Create(scenario.LocalPath, baseRows);
    TestWorkbookFactory.Create(scenario.RemotePath, new[]
    {
        new[] { "1", "key-a", "old-a" },
        new[] { "1", "key-b", "remote-b" }
    });
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        "{\"keyOverrides\":{\"TbTest\":[\"A\"]}}",
        new UTF8Encoding(false));

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal("A", session.KeyName);
    True(session.CanSave);
    session.Save();

    var sheet = new OpenXmlWorkbookReader().Read(scenario.OutputPath).Sheets.Single();
    Equal("remote-b", sheet.GetCell("D5")!.Payload.RawValue);
}

static void IgnoredFieldsPreserveLocal(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "ignored-fields"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "base-note" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local", "local-note" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "remote-note" } });
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        "{\"ignoredFields\":{\"TbTest\":[\"B\"]}}",
        new UTF8Encoding(false));

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.Conflicts.Count);
    Equal("A", session.Conflicts[0].Conflict.FieldName);
    var preview = session.Comparison.CreateTable(MergeGridSide.Merged).Rows.Single(row => row.RecordKey == "1");
    Equal("local-note", preview.Cells[3].DisplayValue);
    session.Conflicts[0].Resolve(MergeChoice.Remote);
    session.Save();

    var row = ReadDataRows(scenario.OutputPath)["1"];
    Equal("remote", row["A"]);
    Equal("local-note", row["B"]);
}

static void IgnoredOnlyModificationDoesNotBlockDeletion(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "ignored-delete"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "same", "base-note" } });
    TestWorkbookFactory.Create(scenario.LocalPath, Array.Empty<string[]>());
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "same", "remote-note" } });
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        "{\"ignoredFields\":{\"TbTest\":[\"B\"]}}",
        new UTF8Encoding(false));

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    True(session.CanSave);
    Equal(0, session.Conflicts.Count);
    session.Save();
    True(ReadDataRows(scenario.OutputPath).Count == 0);
}

static void LogicalTableUniquenessAcceptsDistinctSiblingKeys(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "logical-uniqueness-success"));
    ConfigureLogicalTableInputs(scenario, "Test.xlsx,Sibling.xlsx");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[]
    {
        new[] { "1", "old", "same" },
        new[] { "2", "remote", "new" }
    });
    TestWorkbookFactory.Create(
        Path.Combine(Path.GetDirectoryName(scenario.OutputPath)!, "Sibling.xlsx"),
        new[] { new[] { "99", "sibling", "same" } });
    WriteMergeConfig(scenario, "{\"validateLogicalTableUniqueness\":true}");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    True(session.CanSave);
    session.Save();
    True(ReadDataRows(scenario.OutputPath).ContainsKey("2"));
}

static void LogicalTableUniquenessReportsRemoteCollision(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "logical-uniqueness-collision"));
    ConfigureLogicalTableInputs(scenario, "Test.xlsx,Sibling.xlsx");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[]
    {
        new[] { "1", "old", "same" },
        new[] { "2", "remote", "new" }
    });
    TestWorkbookFactory.Create(
        Path.Combine(Path.GetDirectoryName(scenario.OutputPath)!, "Sibling.xlsx"),
        new[] { new[] { "2", "sibling", "same" } });
    WriteMergeConfig(scenario, "{\"validateLogicalTableUniqueness\":true}");
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.UnsafeWorkbook, exitCode, error);
    var message = error.ToString();
    True(message.Contains("REMOTE", StringComparison.Ordinal));
    True(message.Contains("Sibling.xlsx", StringComparison.Ordinal));
    True(message.Contains("第 4 行", StringComparison.Ordinal));
    True(message.Contains("键 2", StringComparison.Ordinal));
    True(!File.Exists(scenario.OutputPath));
}

static void InactiveSiblingPathsAreExcluded(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "inactive-sibling"));
    ConfigureLogicalTableInputs(scenario, "Test.xlsx,inactive/Sibling.xlsx");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    WriteMergeConfig(
        scenario,
        "{\"inactivePaths\":[\"ConfigLuban/Datas/inactive/**\"],\"validateLogicalTableUniqueness\":true}");

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    True(session.CanSave);
    session.Save();
    Equal("remote", ReadDataRows(scenario.OutputPath)["1"]["A"]);
}

static void InactiveCurrentWorkbookIsRejected(string testRoot)
{
    var original = CreateScenario(Path.Combine(testRoot, "inactive-current"));
    var scenario = original with
    {
        OutputPath = Path.Combine(
            Path.GetDirectoryName(original.OutputPath)!,
            "inactive",
            "Test.xlsx")
    };
    ConfigureLogicalTableInputs(scenario, "inactive/Test.xlsx");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    WriteMergeConfig(scenario, "{\"inactivePaths\":[\"ConfigLuban/Datas/inactive/**\"]}");
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.UnsafeWorkbook, exitCode, error);
    True(error.ToString().Contains("inactivePaths", StringComparison.Ordinal));
    True(!File.Exists(scenario.OutputPath));
}

static void ConfigureLogicalTableInputs(TestScenario scenario, string inputs)
{
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "Datas", "__tables__.csv"),
        "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
        "##,说明,,,,,,\n" +
        $",TbTest,Test,TRUE,\"{inputs}\",Id,map,c\n",
        new UTF8Encoding(false));
}

static void WriteMergeConfig(TestScenario scenario, string json) =>
    File.WriteAllText(
        Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "luban-excel-merge.json"),
        json,
        new UTF8Encoding(false));

static void ProjectValidationSuccessIsReported(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "validation-success"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    CreateValidationCommand(scenario.RepositoryRoot);
    var validator = new FakeProjectValidator();
    var coordinator = new LubanMergeCoordinator(projectValidator: validator);
    var arguments = CreateArguments(scenario).Append("--validate").ToArray();
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(arguments, output, error, coordinator);

    EqualExit(ExitCodes.Success, exitCode, error);
    Equal(1, validator.CallCount);
    True(output.ToString().Contains("项目快速校验：通过", StringComparison.Ordinal));
    Equal("remote", ReadDataRows(scenario.OutputPath)["1"]["A"]);
}

static void ProjectValidationFailureRestoresMerged(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "validation-rollback-existing"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "99", "existing", "output" } });
    var before = HashFile(scenario.OutputPath);
    CreateValidationCommand(scenario.RepositoryRoot);
    var coordinator = new LubanMergeCoordinator(
        projectValidator: new FakeProjectValidator(shouldFail: true));
    var arguments = CreateArguments(scenario).Append("--validate").ToArray();
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(arguments, output, error, coordinator);

    EqualExit(ExitCodes.ProjectValidationFailed, exitCode, error);
    Equal(before, HashFile(scenario.OutputPath));
    True(error.ToString().Contains("fixture validation failed", StringComparison.Ordinal));

    var noOutput = CreateScenario(Path.Combine(testRoot, "validation-rollback-new"));
    TestWorkbookFactory.Create(noOutput.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(noOutput.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(noOutput.RemotePath, new[] { new[] { "1", "remote", "same" } });
    CreateValidationCommand(noOutput.RepositoryRoot);
    var noOutputArguments = CreateArguments(noOutput).Append("--validate").ToArray();
    using var noOutputWriter = new StringWriter();
    using var noOutputError = new StringWriter();

    var noOutputExit = CliApplication.Run(noOutputArguments, noOutputWriter, noOutputError, coordinator);

    EqualExit(ExitCodes.ProjectValidationFailed, noOutputExit, noOutputError);
    True(!File.Exists(noOutput.OutputPath));
}

static void ProjectValidationRunnerHandlesBatchExitCodes(string testRoot)
{
    var root = Path.Combine(testRoot, "validation-runner");
    Directory.CreateDirectory(root);
    var successPath = Path.Combine(root, "success.cmd");
    File.WriteAllText(successPath, "@echo off\r\nset /p answer=\r\nexit /b 0\r\n", Encoding.ASCII);
    var failurePath = Path.Combine(root, "failure.cmd");
    File.WriteAllText(failurePath, "@echo off\r\necho validation-detail\r\nexit /b 7\r\n", Encoding.ASCII);
    var runner = new ProjectValidationRunner();

    runner.Validate(successPath, root, TimeSpan.FromSeconds(10));
    try
    {
        runner.Validate(failurePath, root, TimeSpan.FromSeconds(10));
        throw new InvalidOperationException("Expected ProjectValidationException.");
    }
    catch (ProjectValidationException exception)
    {
        True(exception.Message.Contains("退出码 7", StringComparison.Ordinal));
        True(exception.Message.Contains("validation-detail", StringComparison.Ordinal));
    }
}

static void FullExportRunsInIsolation(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "full-export-success"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    var markerPath = CreateFullExportFixture(scenario, exitCode: 0);
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.Success, exitCode, error);
    Equal("original", File.ReadAllText(markerPath).Trim());
    Equal("remote", ReadDataRows(scenario.OutputPath)["1"]["A"]);
    True(output.ToString().Contains("隔离副本中通过", StringComparison.Ordinal));
}

static void FullExportFailureRestoresMerged(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "full-export-failure"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "99", "existing", "output" } });
    var outputHash = HashFile(scenario.OutputPath);
    var markerPath = CreateFullExportFixture(scenario, exitCode: 7);
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);

    EqualExit(ExitCodes.ProjectValidationFailed, exitCode, error);
    Equal(outputHash, HashFile(scenario.OutputPath));
    Equal("original", File.ReadAllText(markerPath).Trim());
    True(error.ToString().Contains("完整导出校验失败", StringComparison.Ordinal));
    True(error.ToString().Contains("退出码 7", StringComparison.Ordinal));
}

static void PendingValidationBackupRestoresExistingMerged(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "recovery-existing"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "99", "original", "merged" } });
    var originalHash = HashFile(scenario.OutputPath);
    string markerPath;
    using (MergeOutputLease.Acquire(scenario.OutputPath))
        markerPath = MergeOutputRecovery.CreateRollbackMarker(scenario.OutputPath, outputExisted: true);
    File.Delete(scenario.OutputPath);
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "88", "interrupted", "candidate" } });
    var temporaryPath = Path.Combine(
        Path.GetDirectoryName(scenario.OutputPath)!,
        $".{Path.GetFileName(scenario.OutputPath)}.stale.tmp.xlsx");
    File.Copy(scenario.OutputPath, temporaryPath);

    new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    Equal(originalHash, HashFile(scenario.OutputPath));
    True(!File.Exists(markerPath));
    True(!File.Exists(temporaryPath));
}

static void PendingValidationMarkerRemovesNewMerged(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "recovery-new"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote", "same" } });
    string markerPath;
    using (MergeOutputLease.Acquire(scenario.OutputPath))
        markerPath = MergeOutputRecovery.CreateRollbackMarker(scenario.OutputPath, outputExisted: false);
    TestWorkbookFactory.Create(scenario.OutputPath, new[] { new[] { "88", "interrupted", "candidate" } });

    new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));

    True(!File.Exists(scenario.OutputPath));
    True(!File.Exists(markerPath));
}

static string CreateFullExportFixture(TestScenario scenario, int exitCode)
{
    var outputRoot = Path.Combine(scenario.RepositoryRoot, "ConfigOutput");
    Directory.CreateDirectory(outputRoot);
    var markerPath = Path.Combine(outputRoot, "marker.txt");
    File.WriteAllText(markerPath, "original\r\n", Encoding.ASCII);
    var commandPath = Path.Combine(scenario.RepositoryRoot, "ConfigLuban", "full-export.cmd");
    File.WriteAllText(
        commandPath,
        $"@echo off\r\necho changed>\"%~dp0..\\ConfigOutput\\marker.txt\"\r\nexit /b {exitCode}\r\n",
        Encoding.ASCII);
    WriteMergeConfig(
        scenario,
        "{\"validation\":{\"fullExportEnabled\":true,\"fullExportCommand\":\"ConfigLuban/full-export.cmd\"}}");
    return markerPath;
}

static void DiagnosticLogOmitsCellContents(string testRoot)
{
    const string sensitiveValue = "SECRET-CELL-CONTENT-92741";
    var scenario = CreateScenario(Path.Combine(testRoot, "diagnostic-success"));
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", sensitiveValue, "same" } });
    var logPath = Path.Combine(testRoot, "logs", "success.jsonl");
    var arguments = CreateArguments(scenario).Concat(new[] { "--log", logPath }).ToArray();
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(arguments, output, error);

    EqualExit(ExitCodes.Success, exitCode, error);
    var text = File.ReadAllText(logPath);
    True(!text.Contains(sensitiveValue, StringComparison.Ordinal));
    var entries = File.ReadLines(logPath).Select(line => JsonSerializer.Deserialize<JsonElement>(line)).ToArray();
    Equal("started", entries[0].GetProperty("event").GetString());
    Equal("completed", entries[^1].GetProperty("event").GetString());
    var timings = entries[^1].GetProperty("details").GetProperty("preparationMilliseconds");
    True(timings.TryGetProperty("total", out _));
    True(timings.TryGetProperty("workbookRead", out _));
    True(timings.TryGetProperty("sheets", out _));
    var hash = entries[0].GetProperty("details").GetProperty("files").GetProperty("remote").GetProperty("sha256").GetString();
    Equal(64, hash!.Length);
}

static void DiagnosticLogRecordsSanitizedException(string testRoot)
{
    const string sensitiveKey = "SECRET-KEY-55182";
    var scenario = CreateScenario(Path.Combine(testRoot, "diagnostic-failure"));
    var duplicateRows = new[]
    {
        new[] { sensitiveKey, "a", "b" },
        new[] { sensitiveKey, "c", "d" }
    };
    TestWorkbookFactory.Create(scenario.BasePath, duplicateRows);
    TestWorkbookFactory.Create(scenario.LocalPath, duplicateRows);
    TestWorkbookFactory.Create(scenario.RemotePath, duplicateRows);
    var logPath = Path.Combine(testRoot, "logs", "failure.jsonl");
    var arguments = CreateArguments(scenario).Concat(new[] { "--log", logPath }).ToArray();
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(arguments, output, error);

    EqualExit(ExitCodes.UnsafeWorkbook, exitCode, error);
    var text = File.ReadAllText(logPath);
    True(!text.Contains(sensitiveKey, StringComparison.Ordinal));
    var exceptionEntry = File.ReadLines(logPath)
        .Select(line => JsonSerializer.Deserialize<JsonElement>(line))
        .Single(entry => entry.GetProperty("event").GetString() == "exception");
    Equal(ExitCodes.UnsafeWorkbook, exceptionEntry.GetProperty("details").GetProperty("exitCode").GetInt32());
    True(exceptionEntry.GetProperty("details").GetProperty("exceptionTypes").GetArrayLength() >= 1);
    True(exceptionEntry.GetProperty("details").GetProperty("stackTrace").GetString()!.Length > 0);
}

static void DiagnosticPackageExcludesWorkbooks(string testRoot)
{
    const string logText = "{\"event\":\"fixture\",\"details\":{}}\n";
    var root = Path.Combine(testRoot, "diagnostic-package");
    Directory.CreateDirectory(root);
    var logPath = Path.Combine(root, "merge.jsonl");
    var workbookPath = Path.Combine(root, "secret.xlsx");
    var packagePath = Path.Combine(root, "diagnostics.zip");
    File.WriteAllText(logPath, logText, new UTF8Encoding(false));
    TestWorkbookFactory.Create(workbookPath, new[] { new[] { "1", "secret", "cell" } });
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(
        new[] { "diagnostic-package", "--log", logPath, "--output", packagePath },
        output,
        error);

    EqualExit(ExitCodes.Success, exitCode, error);
    using var archive = ZipFile.OpenRead(packagePath);
    var entries = archive.Entries.Select(entry => entry.FullName).Order(StringComparer.Ordinal).ToArray();
    True(entries.SequenceEqual(new[] { "diagnostics.jsonl", "manifest.json" }, StringComparer.Ordinal));
    True(entries.All(entry => !entry.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)));
    using var manifestReader = new StreamReader(archive.GetEntry("manifest.json")!.Open());
    var manifest = JsonSerializer.Deserialize<JsonElement>(manifestReader.ReadToEnd());
    True(!manifest.GetProperty("includesWorkbooks").GetBoolean());
    True(output.ToString().Contains("不包含原始工作簿", StringComparison.Ordinal));
}

static void DiagnosticPackageRejectsDisguisedWorkbook(string testRoot)
{
    var root = Path.Combine(testRoot, "diagnostic-package-disguised-workbook");
    Directory.CreateDirectory(root);
    var workbookPath = Path.Combine(root, "secret.xlsx");
    var disguisedLogPath = Path.Combine(root, "secret.jsonl");
    var packagePath = Path.Combine(root, "diagnostics.zip");
    TestWorkbookFactory.Create(workbookPath, new[] { new[] { "1", "secret", "cell" } });
    File.Copy(workbookPath, disguisedLogPath);
    using var output = new StringWriter();
    using var error = new StringWriter();

    var exitCode = CliApplication.Run(
        new[] { "diagnostic-package", "--log", disguisedLogPath, "--output", packagePath },
        output,
        error);

    EqualExit(ExitCodes.InvalidInput, exitCode, error);
    True(!File.Exists(packagePath));
    True(error.ToString().Contains("JSON Lines", StringComparison.Ordinal));
}

static void ExceptionTypesMapToExitCodes()
{
    Equal(ExitCodes.InvalidInput, ExitCodes.ForException(new MergeInputException("fixture")));
    Equal(ExitCodes.UnsafeWorkbook, ExitCodes.ForException(new UnsafeWorkbookException("fixture")));
    Equal(ExitCodes.WriteValidationFailed, ExitCodes.ForException(new WorkbookWriteException("fixture")));
    Equal(ExitCodes.ProjectValidationFailed, ExitCodes.ForException(new ProjectValidationException("fixture")));
    Equal(ExitCodes.InternalError, ExitCodes.ForException(new Exception("fixture")));
}

static void CreateValidationCommand(string repositoryRoot)
{
    var configRoot = Path.Combine(repositoryRoot, "ConfigLuban");
    Directory.CreateDirectory(configRoot);
    File.WriteAllText(Path.Combine(configRoot, "check.bat"), "@exit /b 0\r\n", Encoding.ASCII);
}

static void DuplicateKeyIsRejected(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "duplicate-key"));
    var duplicateRows = new[]
    {
        new[] { "1", "a", "b" },
        new[] { "1", "c", "d" }
    };
    TestWorkbookFactory.Create(scenario.BasePath, duplicateRows);
    TestWorkbookFactory.Create(scenario.LocalPath, duplicateRows);
    TestWorkbookFactory.Create(scenario.RemotePath, duplicateRows);

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.UnsafeWorkbook, exitCode, error);
    True(!File.Exists(scenario.OutputPath));
    True(error.ToString().Contains("重复", StringComparison.Ordinal));
}

static void DivergentCompositeKeyEditsConflict(string testRoot)
{
    var scenario = CreateScenario(Path.Combine(testRoot, "composite-key-conflict"), "Id+A");
    TestWorkbookFactory.Create(scenario.BasePath, new[] { new[] { "1", "old", "same" } });
    TestWorkbookFactory.Create(scenario.LocalPath, new[] { new[] { "1", "local-key", "local-value" } });
    TestWorkbookFactory.Create(scenario.RemotePath, new[] { new[] { "1", "remote-key", "remote-value" } });

    var session = new LubanMergeCoordinator().Prepare(CommandLineParser.Parse(CreateArguments(scenario)));
    Equal(1, session.RemainingConflicts);
    session.Conflicts[0].Resolve(MergeChoice.Remote);
    var mergedRow = session.Comparison.CreateTable(MergeGridSide.Merged).Rows.Single(row => row.RecordKey == "1 | old");
    Equal("remote-key", mergedRow.Cells[2].DisplayValue);
    Equal("remote-value", mergedRow.Cells[3].DisplayValue);

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.UnresolvedConflicts, exitCode, error);
    True(!File.Exists(scenario.OutputPath));
    True(error.ToString().Contains("主键字段 A", StringComparison.Ordinal));
}

static void RealAccountMerge(string testRoot, string accountPath, string tablesPath)
{
    var root = Path.Combine(testRoot, "real-account");
    var repository = CreateRepository(root);
    var dataRoot = Path.Combine(repository, "ConfigLuban", "Datas");
    Directory.CreateDirectory(dataRoot);
    File.Copy(tablesPath, Path.Combine(dataRoot, "__tables__.csv"));
    var versions = Path.Combine(root, "versions");
    Directory.CreateDirectory(versions);
    var basePath = Path.Combine(versions, "base.xlsx");
    var localPath = Path.Combine(versions, "local.xlsx");
    var remotePath = Path.Combine(versions, "remote.xlsx");
    var outputPath = Path.Combine(dataRoot, "AccountLv.xlsx");
    File.Copy(accountPath, basePath);
    var sourceHash = HashFile(accountPath);
    var sourceSnapshot = new OpenXmlWorkbookReader().Read(basePath);
    var sheetName = sourceSnapshot.Sheets.Single().Name;
    new AtomicWorkbookSaver().Save(
        basePath,
        localPath,
        new WorkbookEdit[] { new SetCellEdit(sheetName, "C5", new Core.CellPayload(Core.CellValueKind.Number, "201")) });
    new AtomicWorkbookSaver().Save(
        basePath,
        remotePath,
        new WorkbookEdit[] { new SetCellEdit(sheetName, "C6", new Core.CellPayload(Core.CellValueKind.Number, "221")) });
    var scenario = new TestScenario(repository, basePath, localPath, remotePath, outputPath);

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.Success, exitCode, error);
    var merged = new OpenXmlWorkbookReader().Read(outputPath).Sheets.Single();
    Equal("201", merged.GetCell("C5")!.Payload.RawValue);
    Equal("221", merged.GetCell("C6")!.Payload.RawValue);
    Equal(sourceHash, HashFile(accountPath));
}

static void RealBattleMerge(string testRoot, string battlePath, string tablesPath)
{
    var root = Path.Combine(testRoot, "real-battle");
    var repository = CreateRepository(root);
    var dataRoot = Path.Combine(repository, "ConfigLuban", "Datas");
    Directory.CreateDirectory(dataRoot);
    File.Copy(tablesPath, Path.Combine(dataRoot, "__tables__.csv"));
    var versions = Path.Combine(root, "versions");
    Directory.CreateDirectory(versions);
    var basePath = Path.Combine(versions, "base.xlsx");
    var localPath = Path.Combine(versions, "local.xlsx");
    var remotePath = Path.Combine(versions, "remote.xlsx");
    var outputPath = Path.Combine(dataRoot, "BattleActionConstConfig.xlsx");
    File.Copy(battlePath, basePath);
    var sourceHash = HashFile(battlePath);
    var sheetName = new OpenXmlWorkbookReader().Read(basePath).Sheets.Single().Name;
    new AtomicWorkbookSaver().Save(
        basePath,
        localPath,
        new WorkbookEdit[] { new SetCellEdit(sheetName, "E6", new Core.CellPayload(Core.CellValueKind.Number, "2")) });
    new AtomicWorkbookSaver().Save(
        basePath,
        remotePath,
        new WorkbookEdit[] { new SetCellEdit(sheetName, "G6", new Core.CellPayload(Core.CellValueKind.Number, "1")) });
    var scenario = new TestScenario(repository, basePath, localPath, remotePath, outputPath);

    using var output = new StringWriter();
    using var error = new StringWriter();
    var exitCode = CliApplication.Run(CreateArguments(scenario), output, error);
    EqualExit(ExitCodes.Success, exitCode, error);
    var merged = new OpenXmlWorkbookReader().Read(outputPath).Sheets.Single();
    Equal("2", merged.GetCell("E6")!.Payload.RawValue);
    Equal("1", merged.GetCell("G6")!.Payload.RawValue);
    True(output.ToString().Contains("主键=Id+SmallGameType", StringComparison.Ordinal));
    Equal(sourceHash, HashFile(battlePath));
}

static string[] CreateArguments(TestScenario scenario, string recalculationMode = "never") =>
[
    "merge",
    "--base", scenario.BasePath,
    "--local", scenario.LocalPath,
    "--remote", scenario.RemotePath,
    "--output", scenario.OutputPath,
    "--repo-root", scenario.RepositoryRoot,
    "--headless",
    "--recalculate-with-excel", recalculationMode
];

static void EqualExit(int expected, int actual, StringWriter error)
{
    if (expected != actual)
        throw new InvalidOperationException($"Expected exit {expected}, got {actual}. Error: {error}");
}

static TestScenario CreateScenario(string root, string index = "Id", string mode = "map")
{
    var repository = CreateRepository(root);
    var dataRoot = Path.Combine(repository, "ConfigLuban", "Datas");
    Directory.CreateDirectory(dataRoot);
    File.WriteAllText(
        Path.Combine(dataRoot, "__tables__.csv"),
        "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
        "##,说明,,,,,,\n" +
        $",TbTest,Test,TRUE,Test.xlsx,{index},{mode},c\n",
        new UTF8Encoding(false));
    var versions = Path.Combine(root, "versions");
    Directory.CreateDirectory(versions);
    return new TestScenario(
        repository,
        Path.Combine(versions, "base.xlsx"),
        Path.Combine(versions, "local.xlsx"),
        Path.Combine(versions, "remote.xlsx"),
        Path.Combine(dataRoot, "Test.xlsx"));
}

static string CreateRepository(string root)
{
    Directory.CreateDirectory(root);
    Directory.CreateDirectory(Path.Combine(root, ".git"));
    return root;
}

static IReadOnlyDictionary<string, IReadOnlyDictionary<string, string?>> ReadDataRows(
    string path,
    string? sheetName = null)
{
    var workbook = new OpenXmlWorkbookReader().Read(path);
    var sheet = sheetName is null ? workbook.Sheets.Single() : workbook.GetSheet(sheetName);
    var result = new Dictionary<string, IReadOnlyDictionary<string, string?>>(StringComparer.Ordinal);
    foreach (var row in sheet.Rows.Where(row => row.RowNumber >= 4))
    {
        var id = row.Cells.FirstOrDefault(cell => cell.ColumnIndex == 1)?.Payload.RawValue;
        if (id is null)
            continue;
        result[id] = new Dictionary<string, string?>
        {
            ["A"] = row.Cells.FirstOrDefault(cell => cell.ColumnIndex == 2)?.Payload.RawValue,
            ["B"] = row.Cells.FirstOrDefault(cell => cell.ColumnIndex == 3)?.Payload.RawValue
        };
    }
    return result;
}

internal sealed record TestScenario(
    string RepositoryRoot,
    string BasePath,
    string LocalPath,
    string RemotePath,
    string OutputPath);

internal sealed record ProcessResult(int ExitCode, string Output, string Error);

internal sealed record GitConflictFixture(
    string RepositoryRoot,
    string DataRoot,
    string TablesPath,
    string WorkbookPath);

internal static class TestWorkbookFactory
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static void Create(string path, IReadOnlyList<string[]> records)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);
        Add(archive, "xl/worksheets/sheet1.xml", CreateSheet(records));
    }

    internal static void CreateWithSchema(
        string path,
        IReadOnlyList<string> fieldNames,
        IReadOnlyList<string> fieldTypes,
        IReadOnlyList<string[]> records)
    {
        if (fieldNames.Count == 0 || fieldNames.Count != fieldTypes.Count)
            throw new ArgumentException("Field names and types must be non-empty and aligned.");
        if (records.Any(record => record.Length != fieldNames.Count))
            throw new ArgumentException("Every record must match the field count.", nameof(records));

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);
        Add(archive, "xl/worksheets/sheet1.xml", CreateSheet(fieldNames, fieldTypes, records));
    }

    internal static void CreateMultiSheet(
        string path,
        params (string Name, IReadOnlyList<string[]> Records)[] sheets)
    {
        if (sheets.Length == 0)
            throw new ArgumentException("At least one sheet is required.", nameof(sheets));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        var contentTypes = new XElement(
            XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Types",
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Default",
                new XAttribute("Extension", "rels"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Default",
                new XAttribute("Extension", "xml"),
                new XAttribute("ContentType", "application/xml")),
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Override",
                new XAttribute("PartName", "/xl/workbook.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")));
        var relationshipNamespace = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var workbookRelationshipNamespace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        var sheetElements = new List<XElement>();
        var relationshipElements = new List<XElement>();
        for (var index = 0; index < sheets.Length; index++)
        {
            var number = index + 1;
            contentTypes.Add(new XElement(
                contentTypes.Name.Namespace + "Override",
                new XAttribute("PartName", $"/xl/worksheets/sheet{number}.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            sheetElements.Add(new XElement(
                Spreadsheet + "sheet",
                new XAttribute("name", sheets[index].Name),
                new XAttribute("sheetId", number),
                new XAttribute(workbookRelationshipNamespace + "id", $"rId{number}")));
            relationshipElements.Add(new XElement(
                relationshipNamespace + "Relationship",
                new XAttribute("Id", $"rId{number}"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", $"worksheets/sheet{number}.xml")));
            Add(archive, $"xl/worksheets/sheet{number}.xml", CreateSheet(sheets[index].Records));
        }

        Add(archive, "[Content_Types].xml", new XDocument(contentTypes).ToString(SaveOptions.DisableFormatting));
        Add(archive, "_rels/.rels", RootRelationships);
        Add(
            archive,
            "xl/workbook.xml",
            new XDocument(new XElement(
                Spreadsheet + "workbook",
                new XAttribute(XNamespace.Xmlns + "r", workbookRelationshipNamespace),
                new XElement(Spreadsheet + "sheets", sheetElements))).ToString(SaveOptions.DisableFormatting));
        Add(
            archive,
            "xl/_rels/workbook.xml.rels",
            new XDocument(new XElement(relationshipNamespace + "Relationships", relationshipElements))
                .ToString(SaveOptions.DisableFormatting));
    }

    internal static void CreateWithFormula(string path, string value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);

        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, "##var", "Id", "A", "B"));
        sheetData.Add(CreateRow(2, "##", "编号", "字段A", "字段B"));
        sheetData.Add(CreateRow(3, "##type", "string", "string", "int"));
        var dataRow = CreateRow(4, string.Empty, "1", value);
        dataRow.Add(new XElement(
            Spreadsheet + "c",
            new XAttribute("r", "D4"),
            new XElement(Spreadsheet + "f", "LEN(C4)"),
            new XElement(Spreadsheet + "v", value.Length)));
        sheetData.Add(dataRow);
        Add(
            archive,
            "xl/worksheets/sheet1.xml",
            new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Spreadsheet + "worksheet", sheetData))
                .ToString(SaveOptions.DisableFormatting));
    }

    private static string CreateSheet(IReadOnlyList<string[]> records)
    {
        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, "##var", "Id", "A", "B"));
        sheetData.Add(CreateRow(2, "##", "编号", "字段A", "字段B"));
        sheetData.Add(CreateRow(3, "##type", "string", "string", "string"));
        for (var index = 0; index < records.Count; index++)
            sheetData.Add(CreateRow(index + 4, string.Empty, records[index][0], records[index][1], records[index][2]));
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Spreadsheet + "worksheet", sheetData)).ToString(SaveOptions.DisableFormatting);
    }

    private static string CreateSheet(
        IReadOnlyList<string> fieldNames,
        IReadOnlyList<string> fieldTypes,
        IReadOnlyList<string[]> records)
    {
        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, new[] { "##var" }.Concat(fieldNames).ToArray()));
        sheetData.Add(CreateRow(2, new[] { "##" }.Concat(fieldNames.Select(name =>
            string.IsNullOrEmpty(name) ? string.Empty : $"字段{name}")).ToArray()));
        sheetData.Add(CreateRow(3, new[] { "##type" }.Concat(fieldTypes).ToArray()));
        for (var index = 0; index < records.Count; index++)
        {
            sheetData.Add(CreateRow(
                index + 4,
                new[] { string.Empty }.Concat(records[index]).ToArray()));
        }
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Spreadsheet + "worksheet", sheetData)).ToString(SaveOptions.DisableFormatting);
    }

    private static XElement CreateRow(int rowNumber, params string[] values)
    {
        var row = new XElement(Spreadsheet + "row", new XAttribute("r", rowNumber));
        for (var column = 0; column < values.Length; column++)
        {
            if (values[column].Length == 0)
                continue;
            row.Add(new XElement(
                Spreadsheet + "c",
                new XAttribute("r", CellReference.Create(rowNumber, column)),
                new XAttribute("t", "inlineStr"),
                new XElement(Spreadsheet + "is", new XElement(Spreadsheet + "t", values[column]))));
        }
        return row;
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private const string ContentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private const string RootRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private const string Workbook = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """;

    private const string WorkbookRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """;
}

internal sealed class FakeWorkbookRecalculator : IWorkbookRecalculator
{
    public string ProviderName => "Test Office";
    public bool IsAvailable => true;
    public int CallCount { get; private set; }

    public void Recalculate(string workbookPath, TimeSpan timeout) => CallCount++;
}

internal sealed class FakeProjectValidator : IProjectValidator
{
    private readonly bool _shouldFail;

    internal FakeProjectValidator(bool shouldFail = false)
    {
        _shouldFail = shouldFail;
    }

    internal int CallCount { get; private set; }

    public void Validate(string commandPath, string repositoryRoot, TimeSpan timeout)
    {
        CallCount++;
        if (_shouldFail)
            throw new ProjectValidationException("fixture validation failed");
    }
}
