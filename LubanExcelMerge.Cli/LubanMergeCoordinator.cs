using System.Text;
using System.Diagnostics;
using LubanExcelMerge.Core;
using LubanExcelMerge.Git;
using LubanExcelMerge.Luban;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Cli;

public sealed record MergeRunResult(
    string LogicalTable,
    string SheetName,
    string KeyName,
    IReadOnlyList<MergeConflict> Conflicts,
    int ChangedCells,
    int AddedRecords,
    int DeletedRecords,
    string? OutputPath,
    WorkbookRecalculationStatus? RecalculationStatus,
    string? RecalculationProvider,
    bool ProjectValidationCompleted,
    bool FullExportValidationCompleted,
    IReadOnlyList<string> IgnoredFields,
    bool LogicalTableUniquenessValidated,
    MergePreparationTimings PreparationTimings,
    string? RecalculationWarning = null)
{
    public bool Succeeded => Conflicts.Count == 0 && OutputPath is not null;
}

public sealed class LubanMergeCoordinator
{
    private const string SingletonRecordKey = "__single_record__";
    private readonly GitLfsInputResolver _lfsResolver;
    private readonly OpenXmlWorkbookReader _workbookReader;
    private readonly AtomicWorkbookSaver _saver;
    private readonly IProjectValidator _projectValidator;
    private readonly IFullExportValidator _fullExportValidator;

    public LubanMergeCoordinator(
        GitLfsInputResolver? lfsResolver = null,
        OpenXmlWorkbookReader? workbookReader = null,
        AtomicWorkbookSaver? saver = null,
        IProjectValidator? projectValidator = null,
        IFullExportValidator? fullExportValidator = null)
    {
        _lfsResolver = lfsResolver ?? new GitLfsInputResolver();
        _workbookReader = workbookReader ?? new OpenXmlWorkbookReader();
        _saver = saver ?? new AtomicWorkbookSaver();
        _projectValidator = projectValidator ?? new ProjectValidationRunner();
        _fullExportValidator = fullExportValidator ?? new IsolatedFullExportValidator();
    }

    public MergeRunResult Merge(MergeCommandOptions options)
    {
        var session = Prepare(options);
        if (!session.CanSave)
        {
            return new MergeRunResult(
                session.LogicalTable,
                session.SheetName,
                session.KeyName,
                session.Conflicts.Select(conflict => conflict.Conflict).ToArray(),
                session.ChangedCells,
                session.AddedRecords,
                session.DeletedRecords,
                null,
                null,
                null,
                false,
                false,
                session.IgnoredFields,
                session.LogicalTableUniquenessValidated,
                session.PreparationTimings);
        }

        var saveResult = session.Save();
        return new MergeRunResult(
            session.LogicalTable,
            session.SheetName,
            session.KeyName,
            Array.Empty<MergeConflict>(),
            session.ChangedCells,
            session.AddedRecords,
            session.DeletedRecords,
            session.OutputPath,
            saveResult.RecalculationStatus,
            saveResult.RecalculationProvider,
            saveResult.ProjectValidationCompleted,
            saveResult.FullExportValidationCompleted,
            session.IgnoredFields,
            session.LogicalTableUniquenessValidated,
            session.PreparationTimings,
            saveResult.RecalculationWarning);
    }

    public PreparedMergeSession Prepare(MergeCommandOptions options)
    {
        var totalStopwatch = Stopwatch.StartNew();
        options = LubanMergeConfigurationLoader.Apply(options);
        var basePath = ResolveInput(options.BasePath, options.RepositoryRoot, "BASE");
        var localPath = ResolveInput(options.LocalPath, options.RepositoryRoot, "LOCAL");
        var remotePath = ResolveInput(options.RemotePath, options.RepositoryRoot, "REMOTE");
        var outputPath = Path.GetFullPath(options.OutputPath);
        EnsureActivePath(outputPath, options.RepositoryRoot, options.InactivePaths);
        if (!string.Equals(Path.GetExtension(outputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("MERGED 输出必须使用 .xlsx 扩展名。");
        if (string.Equals(outputPath, localPath, StringComparison.OrdinalIgnoreCase))
            throw new MergeInputException("MERGED 输出不能与解析后的 LOCAL 输入是同一个文件。");

        using (MergeOutputLease.Acquire(outputPath))
            MergeOutputRecovery.RecoverPending(outputPath);

        var metadata = MetadataDiscovery.Discover(options);
        LogicalTableDefinition logicalTable;
        LogicalTableCatalog catalog;
        try
        {
            using var tablesReader = new StreamReader(metadata.TablesPath, detectEncodingFromByteOrderMarks: true);
            catalog = LogicalTableCatalog.Parse(tablesReader);
            logicalTable = MatchLogicalTable(catalog, metadata.DataRoot, outputPath);
        }

        catch (MergeInputException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException or FormatException)
        {
            throw new MergeInputException("读取 __tables__.csv 失败。", exception);
        }

        ValidateConfiguredLogicalTables(catalog, options);
        ValidateMode(logicalTable);
        WorkbookSnapshot baseWorkbook;
        WorkbookSnapshot localWorkbook;
        WorkbookSnapshot remoteWorkbook;
        long baseReadMilliseconds;
        long localReadMilliseconds;
        long remoteReadMilliseconds;
        try
        {
            baseWorkbook = ReadWorkbookTimed(basePath, out baseReadMilliseconds);
            localWorkbook = ReadWorkbookTimed(localPath, out localReadMilliseconds);
            remoteWorkbook = ReadWorkbookTimed(remotePath, out remoteReadMilliseconds);
        }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException)
        {
            throw new UnsafeWorkbookException("读取工作簿失败或工作簿格式不安全。", exception);
        }
        catch (IOException exception)
        {
            throw new MergeInputException("读取 BASE、LOCAL 或 REMOTE 失败。", exception);
        }

        var configuredKey = GetConfiguredFields(options.KeyOverrides, logicalTable.FullName);
        var ignoredFields = GetConfiguredFields(options.IgnoredFields, logicalTable.FullName) ?? Array.Empty<string>();
        var sheetPreparationStopwatch = Stopwatch.StartNew();
        var sheetSets = SelectSheets(baseWorkbook, localWorkbook, remoteWorkbook);
        if (options.ValidateLogicalTableUniqueness && sheetSets.Count > 1)
        {
            throw new UnsafeWorkbookException(
                "多工作表工作簿暂不支持全逻辑表唯一性扫描；请关闭 validateLogicalTableUniqueness 后逐工作表合并。");
        }
        var preparedSheets = sheetSets.Select(sheets => PrepareSheet(
            sheets,
            basePath,
            localPath,
            remotePath,
            outputPath,
            metadata,
            logicalTable,
            configuredKey,
            ignoredFields,
            options)).ToArray();
        sheetPreparationStopwatch.Stop();
        totalStopwatch.Stop();
        var preparationTimings = new MergePreparationTimings(
            baseReadMilliseconds,
            localReadMilliseconds,
            remoteReadMilliseconds,
            sheetPreparationStopwatch.ElapsedMilliseconds,
            totalStopwatch.ElapsedMilliseconds);
        return new PreparedMergeSession(
            logicalTable.FullName,
            localPath,
            outputPath,
            ignoredFields,
            options.ValidateLogicalTableUniqueness,
            preparedSheets,
            localWorkbook.Sheets.Sum(sheet => sheet.FormulaCount),
            ParseRecalculationMode(options.RecalculateWithExcel ?? "never"),
            _saver,
            options.RepositoryRoot,
            options.ProjectValidationEnabled,
            options.ProjectValidationCommand,
            _projectValidator,
            options.FullExportValidationEnabled,
            options.FullExportValidationCommand,
            _fullExportValidator,
            preparationTimings);
    }

    private WorkbookSnapshot ReadWorkbookTimed(string path, out long elapsedMilliseconds)
    {
        var stopwatch = Stopwatch.StartNew();
        try
        {
            return _workbookReader.Read(path);
        }
        finally
        {
            stopwatch.Stop();
            elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
        }
    }

    private PreparedSheetMerge PrepareSheet(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        string basePath,
        string localPath,
        string remotePath,
        string outputPath,
        MetadataLocation metadata,
        LogicalTableDefinition logicalTable,
        IReadOnlyList<string>? configuredKey,
        IReadOnlyList<string> ignoredFields,
        MergeCommandOptions options)
    {
        try
        {
            var schemas = ParseSchemas(sheets.Base, sheets.Local, sheets.Remote);
            var schemaMerge = CreateSchemaMergePlan(schemas, sheets);
            var metadataPlan = CreateMetadataPlan(
                sheets,
                schemas,
                schemaMerge,
                sheets.Local.Name);
            var dataRows = new[]
            {
                GetDataRows(sheets.Base, schemas.Base),
                GetDataRows(sheets.Local, schemas.Local),
                GetDataRows(sheets.Remote, schemas.Remote)
            };
            var recordSources = new[]
            {
                CreateRecordSources(sheets.Base, schemaMerge, SchemaSide.Base, basePath, dataRows[0]),
                CreateRecordSources(sheets.Local, schemaMerge, SchemaSide.Local, localPath, dataRows[1]),
                CreateRecordSources(sheets.Remote, schemaMerge, SchemaSide.Remote, remotePath, dataRows[2])
            };
            var isSingleton = string.Equals(logicalTable.Mode, "one", StringComparison.OrdinalIgnoreCase);
            RecordKeyDefinition keyDefinition;
            if (isSingleton)
            {
                if (configuredKey is not null)
                {
                    throw new MergeInputException(
                        $"配置文件 {options.LoadedConfigPath ?? "<运行参数>"} 的 " +
                        $"$.keyOverrides['{logicalTable.FullName}'] 不适用于 mode=one 单例逻辑表。");
                }
                ValidateSingletonRows(sheets.Local.Name, dataRows);
                keyDefinition = new RecordKeyDefinition(new[] { SingletonRecordKey });
            }
            else
            {
                var keyCandidates = (configuredKey is null
                        ? logicalTable.DeclaredIndexes
                        : new[] { new RecordKeyDefinition(configuredKey) })
                    .Select(candidate => new RecordKeyDefinition(candidate.FieldNames
                        .Select(fieldName => MapFieldName(schemaMerge, fieldName))))
                    .ToArray();
                var keySelection = PrimaryKeySelector.Select(
                    schemaMerge.TargetSchema,
                    keyCandidates,
                    recordSources[0],
                    recordSources[1],
                    recordSources[2]);
                if (keySelection.Selected is null)
                {
                    var message = FormatKeyFailures(keySelection.Attempts);
                    if (configuredKey is not null)
                    {
                        throw new MergeInputException(
                            $"配置文件 {options.LoadedConfigPath ?? "<运行参数>"} 的 " +
                            $"$.keyOverrides['{logicalTable.FullName}'] 无效：{message}");
                    }
                    throw new UnsafeWorkbookException(message);
                }

                keyDefinition = keySelection.Selected;
            }
            var alignedIgnoredFields = ignoredFields
                .Select(fieldName => MapFieldName(schemaMerge, fieldName))
                .ToArray();
            ValidateIgnoredFields(
                schemaMerge.TargetSchema,
                keyDefinition,
                alignedIgnoredFields,
                options.LoadedConfigPath);
            if (options.ValidateLogicalTableUniqueness && !isSingleton)
            {
                ValidateLogicalTableUniqueness(
                    metadata,
                    logicalTable,
                    outputPath,
                    schemaMerge.TargetSchema,
                    keyDefinition,
                    recordSources,
                    options.RepositoryRoot,
                    options.InactivePaths);
            }
            var ignoredFieldSet = alignedIgnoredFields.ToHashSet(StringComparer.Ordinal);
            var datasets = new[]
            {
                CreateDataset(schemaMerge, SchemaSide.Base, keyDefinition, ignoredFieldSet, dataRows[0]),
                CreateDataset(schemaMerge, SchemaSide.Local, keyDefinition, ignoredFieldSet, dataRows[1]),
                CreateDataset(schemaMerge, SchemaSide.Remote, keyDefinition, ignoredFieldSet, dataRows[2])
            };
            datasets[1] = NormalizeStructurallyDeletedFieldData(
                datasets[0], datasets[1], schemaMerge, SchemaSide.Local);
            datasets[2] = NormalizeStructurallyDeletedFieldData(
                datasets[0], datasets[2], schemaMerge, SchemaSide.Remote);
            var plan = CreateEdits(
                datasets[0],
                datasets[1],
                datasets[2],
                schemaMerge.TargetSchema,
                keyDefinition,
                sheets.Local.Name,
                ignoredFieldSet);
            var conflicts = schemaMerge.Conflicts
                .Concat(metadataPlan.Conflicts)
                .Concat(plan.Conflicts)
                .ToArray();
            var automaticEdits = metadataPlan.Edits.Concat(plan.Edits).ToArray();
            var comparison = CreateComparison(
                sheets,
                schemas,
                schemaMerge,
                datasets[0],
                datasets[1],
                datasets[2],
                conflicts,
                ignoredFieldSet);
            return new PreparedSheetMerge(
                sheets.Local.Name,
                keyDefinition.DisplayName,
                automaticEdits,
                conflicts,
                comparison,
                plan.ChangedCells,
                plan.AddedRecords,
                plan.DeletedRecords,
                schemaMerge.Conflicts.Count + metadataPlan.Conflicts.Count,
                schemaMerge.StructuralChanges,
                () => CreateFinalEdits(
                    sheets,
                    schemaMerge,
                    automaticEdits,
                    conflicts));
        }
        catch (MergeInputException exception)
        {
            throw new MergeInputException($"工作表 {sheets.Local.Name}：{exception.Message}", exception);
        }
        catch (UnsafeWorkbookException exception)
        {
            throw new UnsafeWorkbookException($"工作表 {sheets.Local.Name}：{exception.Message}", exception);
        }
    }

    private string ResolveInput(string path, string repositoryRoot, string label)
    {
        var absolute = Path.GetFullPath(path);
        if (!File.Exists(absolute))
            throw new MergeInputException($"{label} 文件不存在：{absolute}。");
        try
        {
            return _lfsResolver.Resolve(absolute, repositoryRoot).ContentPath;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or DirectoryNotFoundException)
        {
            throw new MergeInputException($"{label} 的 Git LFS 内容不可用：{exception.Message}", exception);
        }
    }

    private static WorkbookRecalculationMode ParseRecalculationMode(string value) => value switch
    {
        "auto" => WorkbookRecalculationMode.Auto,
        "always" => WorkbookRecalculationMode.Always,
        "never" => WorkbookRecalculationMode.Never,
        _ => throw new MergeInputException("--recalculate-with-excel 必须是 auto、always 或 never。")
    };

    private static void EnsureActivePath(
        string outputPath,
        string repositoryRoot,
        IReadOnlyList<string>? inactivePaths)
    {
        var relativePath = Path.GetRelativePath(Path.GetFullPath(repositoryRoot), outputPath);
        if (PathPatternMatcher.IsMatch(relativePath, inactivePaths))
        {
            throw new UnsafeWorkbookException(
                $"MERGED 位于 inactivePaths 排除范围内：{relativePath}。工具不会自动合并停用目录中的工作簿。");
        }
    }

    private static LogicalTableDefinition MatchLogicalTable(
        LogicalTableCatalog catalog,
        string dataRoot,
        string outputPath)
    {
        var relativePath = Path.GetRelativePath(dataRoot, outputPath);
        if (relativePath.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) || Path.IsPathRooted(relativePath))
            relativePath = Path.GetFileName(outputPath);
        var matches = catalog.MatchInput(relativePath);
        if (matches.Count == 0)
            throw new UnsafeWorkbookException($"__tables__.csv 中没有匹配输入 {relativePath} 的逻辑表。");
        if (matches.Count > 1)
            throw new UnsafeWorkbookException($"输入 {relativePath} 同时匹配多个逻辑表：{string.Join(", ", matches.Select(table => table.FullName))}。");
        return matches[0];
    }

    private static void ValidateConfiguredLogicalTables(
        LogicalTableCatalog catalog,
        MergeCommandOptions options)
    {
        var knownTables = catalog.Tables.Select(table => table.FullName).ToHashSet(StringComparer.Ordinal);
        foreach (var mapping in new[]
                 {
                     (Name: "keyOverrides", Values: options.KeyOverrides),
                     (Name: "ignoredFields", Values: options.IgnoredFields)
                 })
        {
            if (mapping.Values is null)
                continue;
            foreach (var tableName in mapping.Values.Keys.Where(tableName => !knownTables.Contains(tableName)))
            {
                throw new MergeInputException(
                    $"配置文件 {options.LoadedConfigPath ?? "<运行参数>"} 的 $.{mapping.Name}['{tableName}'] " +
                    "未匹配 __tables__.csv 中的逻辑表，逻辑表名区分大小写。");
            }
        }
    }

    private static string[]? GetConfiguredFields(
        IReadOnlyDictionary<string, string[]>? mappings,
        string logicalTable) =>
        mappings is not null && mappings.TryGetValue(logicalTable, out var fields) ? fields : null;

    private static void ValidateIgnoredFields(
        LubanSchema schema,
        RecordKeyDefinition keyDefinition,
        IReadOnlyList<string> ignoredFields,
        string? configPath)
    {
        foreach (var fieldName in ignoredFields)
        {
            if (schema.FindField(fieldName) is null)
            {
                throw new MergeInputException(
                    $"配置文件 {configPath ?? "<运行参数>"} 的 ignoredFields 字段 {fieldName} 不存在，字段名区分大小写。");
            }
            if (keyDefinition.FieldNames.Contains(fieldName, StringComparer.Ordinal))
            {
                throw new MergeInputException(
                    $"配置文件 {configPath ?? "<运行参数>"} 不能忽略主键字段 {fieldName}。");
            }
        }
    }

    private void ValidateLogicalTableUniqueness(
        MetadataLocation metadata,
        LogicalTableDefinition logicalTable,
        string outputPath,
        LubanSchema currentSchema,
        RecordKeyDefinition keyDefinition,
        IReadOnlyList<LubanRecordSource>[] currentSources,
        string repositoryRoot,
        IReadOnlyList<string>? inactivePaths)
    {
        if (logicalTable.InputFiles.Count <= 1)
            return;

        var currentInput = FindCurrentLogicalInput(logicalTable, metadata.DataRoot, outputPath);
        var siblingSources = new List<LubanRecordSource>();
        foreach (var input in logicalTable.InputFiles.Where(input =>
                     !string.Equals(NormalizeLogicalInput(input), currentInput, StringComparison.OrdinalIgnoreCase)))
        {
            var extension = Path.GetExtension(input);
            if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new UnsafeWorkbookException(
                    $"逻辑表 {logicalTable.FullName} 的兄弟输入 {input} 不是 .xlsx，当前版本不能完整执行跨文件唯一性校验。");
            }

            var siblingPath = Path.GetFullPath(input.Replace('/', Path.DirectorySeparatorChar), metadata.DataRoot);
            var repositoryRelativePath = Path.GetRelativePath(Path.GetFullPath(repositoryRoot), siblingPath);
            if (PathPatternMatcher.IsMatch(repositoryRelativePath, inactivePaths))
                continue;
            if (!File.Exists(siblingPath))
                throw new MergeInputException($"逻辑表 {logicalTable.FullName} 的兄弟工作簿不存在：{siblingPath}。");

            WorkbookSnapshot workbook;
            try
            {
                var resolvedPath = ResolveInput(siblingPath, repositoryRoot, $"逻辑表兄弟文件 {input}");
                workbook = _workbookReader.Read(resolvedPath);
            }
            catch (MergeInputException)
            {
                throw;
            }
            catch (Exception exception) when (exception is IOException or InvalidDataException or NotSupportedException)
            {
                throw new UnsafeWorkbookException($"读取逻辑表兄弟工作簿失败：{siblingPath}。", exception);
            }

            if (workbook.Sheets.Count != 1)
            {
                throw new UnsafeWorkbookException(
                    $"逻辑表兄弟工作簿 {siblingPath} 包含 {workbook.Sheets.Count} 个工作表，当前版本不能安全扫描。");
            }

            var sheet = workbook.Sheets[0];
            LubanSchema siblingSchema;
            try
            {
                siblingSchema = ParseSchema(sheet);
            }
            catch (FormatException exception)
            {
                throw new UnsafeWorkbookException($"无法识别逻辑表兄弟工作簿的 Luban 表头：{siblingPath}。", exception);
            }
            if (siblingSchema.IsRestricted)
            {
                throw new UnsafeWorkbookException(
                    $"逻辑表兄弟工作簿 {siblingPath} 结构受限：{string.Join("；", siblingSchema.Restrictions)}");
            }
            foreach (var fieldName in keyDefinition.FieldNames)
            {
                if (siblingSchema.FindField(fieldName) is null)
                {
                    throw new UnsafeWorkbookException(
                        $"逻辑表兄弟工作簿 {siblingPath} 缺少主键字段 {fieldName}，字段名区分大小写。");
                }
            }
            siblingSources.AddRange(CreateRecordSources(
                sheet,
                siblingSchema,
                siblingPath,
                GetDataRows(sheet, siblingSchema)));
        }

        var variants = new[] { "BASE", "LOCAL", "REMOTE" };
        for (var index = 0; index < currentSources.Length; index++)
        {
            var combined = currentSources[index].Concat(siblingSources).ToArray();
            var result = RecordKeyValidator.Validate(currentSchema, keyDefinition, combined);
            if (!result.IsValid)
            {
                var details = result.Issues.Take(20).Select(issue =>
                {
                    var workbook = Path.GetRelativePath(metadata.DataRoot, issue.WorkbookPath);
                    var key = string.IsNullOrEmpty(issue.KeyValue) ? "<空>" : issue.KeyValue;
                    return $"{variants[index]}：{workbook} / {issue.SheetName} / 第 {issue.RowNumber} 行 / 键 {key}：{issue.Message}";
                });
                throw new UnsafeWorkbookException(
                    $"逻辑表 {logicalTable.FullName} 的跨文件主键 {keyDefinition.DisplayName} 不唯一或不完整：" +
                    string.Join("；", details));
            }
        }
    }

    private static string FindCurrentLogicalInput(
        LogicalTableDefinition logicalTable,
        string dataRoot,
        string outputPath)
    {
        var relativeOutput = NormalizeLogicalInput(Path.GetRelativePath(dataRoot, outputPath));
        var exact = logicalTable.InputFiles.FirstOrDefault(input =>
            string.Equals(NormalizeLogicalInput(input), relativeOutput, StringComparison.OrdinalIgnoreCase));
        if (exact is not null)
            return NormalizeLogicalInput(exact);

        var outputName = Path.GetFileName(outputPath);
        var fileNameMatches = logicalTable.InputFiles.Where(input =>
            string.Equals(Path.GetFileName(input), outputName, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (fileNameMatches.Length == 1)
            return NormalizeLogicalInput(fileNameMatches[0]);

        throw new UnsafeWorkbookException(
            $"无法确定逻辑表 {logicalTable.FullName} 中与 MERGED 对应的当前输入文件，不能执行跨文件唯一性校验。");
    }

    private static string NormalizeLogicalInput(string path) =>
        path.Trim().Replace('\\', '/').TrimStart('/');

    private static void ValidateMode(LogicalTableDefinition table)
    {
        if (!string.IsNullOrEmpty(table.Mode) &&
            !string.Equals(table.Mode, "map", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(table.Mode, "one", StringComparison.OrdinalIgnoreCase))
            throw new UnsafeWorkbookException($"逻辑表 {table.FullName} 的模式 {table.Mode} 尚不能安全地在无界面模式下合并。");
    }

    private static void ValidateSingletonRows(
        string sheetName,
        IReadOnlyList<IReadOnlyList<OpenXmlRowSnapshot>> dataRows)
    {
        foreach (var (side, rows) in dataRows.Select((rows, index) => (new[] { "BASE", "LOCAL", "REMOTE" }[index], rows)))
        {
            if (rows.Count != 1)
            {
                throw new UnsafeWorkbookException(
                    $"工作表 {sheetName} 的 mode=one 表在 {side} 中必须恰好有 1 条数据记录，实际为 {rows.Count} 条。");
            }
        }
    }

    private static IReadOnlyList<(SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote)> SelectSheets(
        WorkbookSnapshot @base,
        WorkbookSnapshot local,
        WorkbookSnapshot remote)
    {
        var baseNames = @base.Sheets.Select(sheet => sheet.Name).ToArray();
        var localNames = local.Sheets.Select(sheet => sheet.Name).ToArray();
        var remoteNames = remote.Sheets.Select(sheet => sheet.Name).ToArray();
        if (baseNames.Length == 0 ||
            !baseNames.SequenceEqual(localNames, StringComparer.Ordinal) ||
            !baseNames.SequenceEqual(remoteNames, StringComparer.Ordinal))
        {
            throw new UnsafeWorkbookException("BASE、LOCAL、REMOTE 的工作表名称、数量或顺序不一致。");
        }
        return Enumerable.Range(0, baseNames.Length)
            .Select(index => (@base.Sheets[index], local.Sheets[index], remote.Sheets[index]))
            .ToArray();
    }

    private static (LubanSchema Base, LubanSchema Local, LubanSchema Remote) ParseSchemas(
        SheetSnapshot @base,
        SheetSnapshot local,
        SheetSnapshot remote)
    {
        try
        {
            return (
                NormalizeSchema(ParseSchema(@base)),
                NormalizeSchema(ParseSchema(local)),
                NormalizeSchema(ParseSchema(remote)));
        }
        catch (FormatException exception)
        {
            throw new UnsafeWorkbookException("无法识别 Luban 表头。", exception);
        }
    }

    private static LubanSchema ParseSchema(SheetSnapshot sheet)
    {
        var rows = sheet.Rows.Select(row =>
        {
            var maximumColumn = row.Cells.Select(cell => cell.ColumnIndex).DefaultIfEmpty(-1).Max();
            var values = new string?[maximumColumn + 1];
            foreach (var cell in row.Cells)
                values[cell.ColumnIndex] = cell.Payload.RawValue;
            return new LubanRawRow(row.RowNumber, values);
        }).ToArray();
        return LubanSchemaParser.Parse(rows);
    }

    private static LubanSchema NormalizeSchema(LubanSchema schema)
    {
        // A commented first data row also starts with ##. It belongs to the
        // data area, not to Luban metadata, and must not shift DataStartRowNumber.
        var metadataRows = schema.MetadataRows
            .Where(row => !IsCommentedDataRow(schema, row))
            .ToArray();
        var dataStart = metadataRows.Length == 0
            ? schema.DataStartRowNumber
            : metadataRows.Max(row => row.RowNumber) + 1;
        return schema with
        {
            MetadataRows = metadataRows,
            DataStartRowNumber = dataStart
        };
    }

    private static bool IsCommentedDataRow(LubanSchema schema, LubanRawRow row)
    {
        var marker = row.Cells.FirstOrDefault(cell => !string.IsNullOrEmpty(cell));
        if (!string.Equals(marker, "##", StringComparison.Ordinal) ||
            row.RowNumber <= schema.PrimaryVariableRowNumber)
        {
            return false;
        }

        // Numeric and boolean fields provide an unambiguous signal that a ## row
        // is a commented record rather than a descriptive metadata row.
        foreach (var field in schema.Fields)
        {
            if (field.TypeName.Length == 0 ||
                field.TypeName.StartsWith("string", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = field.ColumnIndex < row.Cells.Count ? row.Cells[field.ColumnIndex] : null;
            if (string.IsNullOrWhiteSpace(value))
                continue;
            if (field.TypeName.StartsWith("bool", StringComparison.OrdinalIgnoreCase) &&
                (value is "0" or "1" or "true" or "false"))
            {
                return true;
            }
            if (field.TypeName.StartsWith("int", StringComparison.OrdinalIgnoreCase) ||
                field.TypeName.StartsWith("long", StringComparison.OrdinalIgnoreCase) ||
                field.TypeName.StartsWith("float", StringComparison.OrdinalIgnoreCase) ||
                field.TypeName.StartsWith("double", StringComparison.OrdinalIgnoreCase) ||
                field.TypeName.StartsWith("decimal", StringComparison.OrdinalIgnoreCase))
            {
                if (decimal.TryParse(value, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static SchemaMergePlan CreateSchemaMergePlan(
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas,
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets)
    {
        if (schemas.Base.IsRestricted || schemas.Local.IsRestricted || schemas.Remote.IsRestricted)
        {
            var restrictions = schemas.Base.Restrictions
                .Concat(schemas.Local.Restrictions)
                .Concat(schemas.Remote.Restrictions)
                .Distinct(StringComparer.Ordinal);
            throw new UnsafeWorkbookException("工作表结构处于受限模式：" + string.Join("；", restrictions));
        }

        var drafts = AlignFields(schemas);
        var occupiedLocalColumns = sheets.Local.Rows
            .SelectMany(row => row.Cells)
            .Where(HasCellContent)
            .Select(cell => cell.ColumnIndex)
            .ToHashSet();
        var usedTargetColumns = schemas.Local.Fields.Select(field => field.ColumnIndex).ToHashSet();
        var nextTargetColumn = occupiedLocalColumns.DefaultIfEmpty(-1).Max() + 1;
        var alignments = new List<AlignedField>();
        var structuralConflicts = new List<ResolvableMergeConflict>();
        var structuralChanges = new List<string>();

        foreach (var draft in drafts)
        {
            var analysisSource = draft.Local ?? draft.Base ?? draft.Remote
                ?? throw new InvalidOperationException("字段对齐缺少三方来源。");
            var targetColumn = draft.Local?.ColumnIndex ?? draft.Base?.ColumnIndex ?? draft.Remote!.ColumnIndex;
            if (occupiedLocalColumns.Contains(targetColumn) || usedTargetColumns.Contains(targetColumn))
            {
                if (draft.Local is null)
                {
                    while (occupiedLocalColumns.Contains(nextTargetColumn) || usedTargetColumns.Contains(nextTargetColumn))
                        nextTargetColumn++;
                    targetColumn = nextTargetColumn++;
                }
            }
            while (usedTargetColumns.Contains(targetColumn) && draft.Local?.ColumnIndex != targetColumn)
            {
                while (occupiedLocalColumns.Contains(nextTargetColumn) || usedTargetColumns.Contains(nextTargetColumn))
                    nextTargetColumn++;
                targetColumn = nextTargetColumn++;
            }
            usedTargetColumns.Add(targetColumn);
            var targetField = analysisSource with { ColumnIndex = targetColumn };
            var decision = MergeFieldStructure(draft.Base, draft.Local, draft.Remote);
            ResolvableMergeConflict? structuralConflict = null;
            if (decision.IsConflict)
            {
                var address = CellReference.Create(schemas.Base.PrimaryVariableRowNumber, targetColumn);
                var conflictKind = draft.Base is not null &&
                                   (draft.Local is null || draft.Remote is null)
                    ? MergeConflictKind.DeleteModify
                    : MergeConflictKind.MetadataChanged;
                var conflict = new MergeConflict(
                    conflictKind,
                    $"字段 {draft.DisplayName}",
                    address,
                    $"字段 {draft.DisplayName} 的列结构在 LOCAL 和 REMOTE 中发生不兼容变化，请选择保留 BASE、LOCAL 或 REMOTE。");
                structuralConflict = CreateResolution(
                    structuralConflicts.Count,
                    conflict,
                    schemas.Base.PrimaryVariableRowNumber,
                    FormatField(draft.Base),
                    FormatField(draft.Local),
                    FormatField(draft.Remote),
                    Array.Empty<WorkbookEdit>(),
                    Array.Empty<WorkbookEdit>(),
                    Array.Empty<WorkbookEdit>());
                structuralConflicts.Add(structuralConflict);
            }

            alignments.Add(new AlignedField(
                draft.Identity,
                targetField,
                draft.Base,
                draft.Local,
                draft.Remote,
                decision,
                structuralConflict));
            AddStructuralChangeSummaries(structuralChanges, draft);
        }

        alignments.Sort((left, right) => left.Target.ColumnIndex.CompareTo(right.Target.ColumnIndex));
        var duplicateTargetNames = alignments
            .GroupBy(alignment => alignment.Target.Name, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToArray();
        if (duplicateTargetNames.Length > 0)
        {
            throw new UnsafeWorkbookException(
                "列结构合并后存在无法唯一对齐的字段名：" + string.Join("、", duplicateTargetNames) + "。");
        }
        var targetSchema = new LubanSchema(
            alignments.Select(alignment => alignment.Target).ToArray(),
            schemas.Local.MetadataRows,
            schemas.Local.PrimaryVariableRowNumber,
            schemas.Local.TypeRowNumber,
            schemas.Local.DataStartRowNumber,
            false,
            Array.Empty<string>());
        var maximumTargetColumn = alignments.Select(alignment => alignment.Target.ColumnIndex).DefaultIfEmpty(-1).Max();
        var maximumRawColumn = sheets.Local.Rows
            .SelectMany(row => row.Cells)
            .Where(HasCellContent)
            .Select(cell => cell.ColumnIndex)
            .DefaultIfEmpty(-1)
            .Max();
        return new SchemaMergePlan(
            targetSchema,
            alignments,
            Math.Max(maximumTargetColumn, maximumRawColumn) + 1,
            structuralConflicts,
            structuralChanges.Distinct(StringComparer.Ordinal).ToArray(),
            schemas);
    }

    private static IReadOnlyList<FieldAlignmentDraft> AlignFields(
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas)
    {
        var localRemaining = schemas.Local.Fields.ToHashSet();
        var remoteRemaining = schemas.Remote.Fields.ToHashSet();
        var drafts = schemas.Base.Fields
            .Select(field => new MutableFieldAlignment(
                $"base:{field.Name}",
                field.Name,
                field,
                TakeByName(localRemaining, field.Name),
                TakeByName(remoteRemaining, field.Name)))
            .ToArray();

        foreach (var draft in drafts.Where(draft => draft.Local is null))
            draft.Local = TakeByColumn(localRemaining, draft.Base!.ColumnIndex);
        foreach (var draft in drafts.Where(draft => draft.Remote is null))
            draft.Remote = TakeByColumn(remoteRemaining, draft.Base!.ColumnIndex);

        var result = drafts.Select(draft => draft.ToImmutable()).ToList();
        foreach (var localField in localRemaining.OrderBy(field => field.ColumnIndex).ToArray())
        {
            localRemaining.Remove(localField);
            var remoteField = TakeByName(remoteRemaining, localField.Name);
            result.Add(new FieldAlignmentDraft(
                $"added:{localField.Name}:{localField.ColumnIndex}",
                localField.Name,
                null,
                localField,
                remoteField));
        }
        foreach (var remoteField in remoteRemaining.OrderBy(field => field.ColumnIndex))
        {
            result.Add(new FieldAlignmentDraft(
                $"added:{remoteField.Name}:{remoteField.ColumnIndex}:remote",
                remoteField.Name,
                null,
                null,
                remoteField));
        }
        return result;
    }

    private static LubanField? TakeByName(ISet<LubanField> fields, string name) =>
        TakeField(fields, field => string.Equals(field.Name, name, StringComparison.Ordinal));

    private static LubanField? TakeByColumn(ISet<LubanField> fields, int columnIndex) =>
        TakeField(fields, field => field.ColumnIndex == columnIndex);

    private static LubanField? TakeField(ISet<LubanField> fields, Func<LubanField, bool> predicate)
    {
        var match = fields.FirstOrDefault(predicate);
        if (match is not null)
            fields.Remove(match);
        return match;
    }

    private static FieldStructureDecision MergeFieldStructure(
        LubanField? @base,
        LubanField? local,
        LubanField? remote)
    {
        if (@base is null)
        {
            if (local is null)
                return new FieldStructureDecision(remote, MergeChoice.Remote, false);
            if (remote is null)
                return new FieldStructureDecision(local, MergeChoice.Local, false);
            return FieldEquals(local, remote)
                ? new FieldStructureDecision(local, MergeChoice.Local, false)
                : new FieldStructureDecision(null, MergeChoice.Local, true);
        }

        if (local is null && remote is null)
            return new FieldStructureDecision(null, MergeChoice.Local, false);
        if (local is null)
            return FieldEquals(@base, remote)
                ? new FieldStructureDecision(null, MergeChoice.Local, false)
                : new FieldStructureDecision(null, MergeChoice.Local, true);
        if (remote is null)
            return FieldEquals(@base, local)
                ? new FieldStructureDecision(null, MergeChoice.Remote, false)
                : new FieldStructureDecision(null, MergeChoice.Local, true);
        if (FieldEquals(local, remote) || FieldEquals(@base, remote))
            return new FieldStructureDecision(local, MergeChoice.Local, false);
        if (FieldEquals(@base, local))
            return new FieldStructureDecision(remote, MergeChoice.Remote, false);
        return new FieldStructureDecision(null, MergeChoice.Local, true);
    }

    private static bool FieldEquals(LubanField? left, LubanField? right) =>
        left == right;

    private static string FormatField(LubanField? field) => field is null
        ? "<已删除>"
        : $"{field.Name} : {field.TypeName} @ {MergeComparison.ColumnName(field.ColumnIndex)}列";

    private static void AddStructuralChangeSummaries(
        ICollection<string> summaries,
        FieldAlignmentDraft draft)
    {
        if (draft.Base is null)
            return;
        foreach (var (side, field) in new[] { ("LOCAL", draft.Local), ("REMOTE", draft.Remote) })
        {
            if (field is null)
            {
                summaries.Add($"{side} 删除了既有字段 {draft.Base.Name}");
                continue;
            }
            if (!string.Equals(field.Name, draft.Base.Name, StringComparison.Ordinal))
                summaries.Add($"{side} 将字段 {draft.Base.Name} 重命名为 {field.Name}");
            if (!string.Equals(field.TypeName, draft.Base.TypeName, StringComparison.Ordinal))
                summaries.Add($"{side} 修改了字段 {draft.Base.Name} 的类型");
            if (field.ColumnIndex != draft.Base.ColumnIndex)
            {
                summaries.Add(
                    $"{side} 将字段 {draft.Base.Name} 从 {MergeComparison.ColumnName(draft.Base.ColumnIndex)}列" +
                    $"移动到 {MergeComparison.ColumnName(field.ColumnIndex)}列");
            }
        }
    }

    private static MetadataMergePlan CreateMetadataPlan(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas,
        SchemaMergePlan schemaMerge,
        string sheetName)
    {
        var edits = new List<WorkbookEdit>();
        var conflicts = new List<ResolvableMergeConflict>();
        var metadataRows = new[] { schemas.Base.MetadataRows, schemas.Local.MetadataRows, schemas.Remote.MetadataRows };
        if (metadataRows.Any(rows => rows.Count != schemas.Local.MetadataRows.Count) ||
            !MetadataShapeEquals(schemas.Base, schemas.Local) ||
            !MetadataShapeEquals(schemas.Remote, schemas.Local))
        {
            var baseValue = FormatMetadataRows(sheets.Base, schemas.Base, schemaMerge, SchemaSide.Base);
            var localValue = FormatMetadataRows(sheets.Local, schemas.Local, schemaMerge, SchemaSide.Local);
            var remoteValue = FormatMetadataRows(sheets.Remote, schemas.Remote, schemaMerge, SchemaSide.Remote);
            var conflict = new MergeConflict(
                MergeConflictKind.MetadataChanged,
                "Luban 元数据行",
                CellReference.Create(1, 0),
                "BASE、LOCAL、REMOTE 的 Luban 元数据行数量或数据起始行不同，请选择要保留的元数据区。");
            conflicts.Add(CreateResolution(
                conflicts.Count,
                conflict,
                1,
                baseValue,
                localValue,
                remoteValue,
                new[] { CreateMetadataReplacement(sheets.Local.Name, schemas.Local, schemaMerge, sheets.Base, schemas.Base, SchemaSide.Base) },
                Array.Empty<WorkbookEdit>(),
                new[] { CreateMetadataReplacement(sheets.Local.Name, schemas.Local, schemaMerge, sheets.Remote, schemas.Remote, SchemaSide.Remote) }));
            return new MetadataMergePlan(edits, conflicts);
        }
        var alignmentByColumn = schemaMerge.Fields.ToDictionary(
            alignment => alignment.Target.ColumnIndex);
        for (var rowNumber = 1; rowNumber <= schemas.Local.MetadataRows.Count; rowNumber++)
        {
            var baseCells = CreateAlignedCellArray(
                FindRow(sheets.Base, rowNumber), schemas.Base, schemaMerge, SchemaSide.Base);
            var localCells = CreateAlignedCellArray(
                FindRow(sheets.Local, rowNumber), schemas.Local, schemaMerge, SchemaSide.Local);
            var remoteCells = CreateAlignedCellArray(
                FindRow(sheets.Remote, rowNumber), schemas.Remote, schemaMerge, SchemaSide.Remote);
            for (var columnIndex = 0; columnIndex < schemaMerge.ColumnCount; columnIndex++)
            {
                var baseCell = baseCells[columnIndex];
                var localCell = localCells[columnIndex];
                var remoteCell = remoteCells[columnIndex];
                if (baseCell.ContentEquals(localCell) && baseCell.ContentEquals(remoteCell))
                    continue;

                var address = CellReference.Create(rowNumber, columnIndex);
                var alignment = alignmentByColumn.GetValueOrDefault(columnIndex);
                if (alignment?.HasStructuralVariation == true ||
                    rowNumber == schemaMerge.TargetSchema.PrimaryVariableRowNumber ||
                    rowNumber == schemaMerge.TargetSchema.TypeRowNumber)
                {
                    continue;
                }
                if (alignment is { Base: null, Local: null, Remote: not null })
                {
                    if (!remoteCell.ContentEquals(localCell))
                    {
                        var sourceRow = FindRow(sheets.Remote, rowNumber);
                        var sourceCell = sourceRow is null
                            ? null
                            : GetCell(sourceRow, alignment.Remote.ColumnIndex);
                        edits.Add(new SetCellEdit(sheetName, address, remoteCell, sourceCell?.StyleIndex));
                    }
                    continue;
                }
                if (alignment is { Base: null, Local: not null, Remote: null })
                    continue;
                if (alignment is { Base: null, Local: not null, Remote: not null } &&
                    localCell.ContentEquals(remoteCell))
                {
                    continue;
                }

                var changeSource = baseCell.ContentEquals(localCell)
                    ? "REMOTE 修改了"
                    : baseCell.ContentEquals(remoteCell)
                        ? "LOCAL 修改了"
                        : localCell.ContentEquals(remoteCell)
                            ? "LOCAL 和 REMOTE 同样修改了"
                            : "LOCAL 和 REMOTE 分别修改了";
                var conflict = new MergeConflict(
                    MergeConflictKind.MetadataChanged,
                    $"元数据行 {rowNumber}",
                    address,
                    $"{changeSource} Luban 元数据 {address}，请重点检查并明确选择保留值。");
                conflicts.Add(CreateResolution(
                    conflicts.Count,
                    conflict,
                    rowNumber,
                    FormatCell(baseCell),
                    FormatCell(localCell),
                    FormatCell(remoteCell),
                    CreateSetCellChoice(sheetName, address, localCell, baseCell),
                    Array.Empty<WorkbookEdit>(),
                    CreateSetCellChoice(sheetName, address, localCell, remoteCell)));
            }
        }
        return new MetadataMergePlan(edits, conflicts);
    }

    private static string FormatMetadataRows(
        SheetSnapshot sheet,
        LubanSchema schema,
        SchemaMergePlan schemaMerge,
        SchemaSide side) =>
        string.Join(" / ", schema.MetadataRows.Select(row =>
            string.Join("|", CreateAlignedCellArray(sheet.GetRow(row.RowNumber), schema, schemaMerge, side)
                .Select(FormatCell))));

    private static bool MetadataShapeEquals(LubanSchema left, LubanSchema right)
    {
        var leftMarkers = left.MetadataRows.Select(row => row.Cells.FirstOrDefault(cell => !string.IsNullOrEmpty(cell)) ?? string.Empty);
        var rightMarkers = right.MetadataRows.Select(row => row.Cells.FirstOrDefault(cell => !string.IsNullOrEmpty(cell)) ?? string.Empty);
        return leftMarkers.SequenceEqual(rightMarkers, StringComparer.Ordinal);
    }

    private static WorkbookEdit CreateMetadataReplacement(
        string sheetName,
        LubanSchema localSchema,
        SchemaMergePlan schemaMerge,
        SheetSnapshot sourceSheet,
        LubanSchema sourceSchema,
        SchemaSide sourceSide) =>
        new ReplaceMetadataRowsEdit(
            sheetName,
            1,
            localSchema.MetadataRows.Count,
            sourceSchema.MetadataRows
                .OrderBy(metadata => metadata.RowNumber)
                .Select(metadata => sourceSheet.GetRow(metadata.RowNumber))
                .Where(row => row is not null)
                .Select(row => new RowWrite(
                    CreateAlignedCellArray(row, sourceSchema, schemaMerge, sourceSide)
                        .Select((payload, column) => new CellWrite(
                            column,
                            payload,
                            row!.GetCell(column)?.StyleIndex))
                        .Where(cell => cell.Payload.Kind != CellValueKind.Blank)
                        .ToArray()))
                .ToArray());

    private static IReadOnlyList<LubanRecordSource> CreateRecordSources(
        SheetSnapshot sheet,
        LubanSchema schema,
        string workbookPath,
        IReadOnlyList<OpenXmlRowSnapshot> dataRows) =>
        dataRows
            .Select(row => new LubanRecordSource(
                workbookPath,
                sheet.Name,
                row.RowNumber,
                schema.Fields.ToDictionary(
                    field => field.Name,
                    field => GetCell(row, field.ColumnIndex)?.Payload.RawValue,
                    StringComparer.Ordinal)))
            .ToArray();

    private static IReadOnlyList<LubanRecordSource> CreateRecordSources(
        SheetSnapshot sheet,
        SchemaMergePlan schemaMerge,
        SchemaSide side,
        string workbookPath,
        IReadOnlyList<OpenXmlRowSnapshot> dataRows) =>
        dataRows
            .Select(row => new LubanRecordSource(
                workbookPath,
                sheet.Name,
                row.RowNumber,
                schemaMerge.Fields.ToDictionary(
                    alignment => alignment.Target.Name,
                    alignment => GetSourceField(alignment, side) is { } field
                        ? GetCell(row, field.ColumnIndex)?.Payload.RawValue
                        : null,
                    StringComparer.Ordinal)))
            .ToArray();

    private static ParsedDataset CreateDataset(
        SchemaMergePlan schemaMerge,
        SchemaSide side,
        RecordKeyDefinition keyDefinition,
        IReadOnlySet<string> ignoredFields,
        IReadOnlyList<OpenXmlRowSnapshot> dataRows)
    {
        var targetSchema = schemaMerge.TargetSchema;
        var records = new List<ParsedRecord>();
        foreach (var row in dataRows)
        {
            var cells = schemaMerge.Fields.ToDictionary(
                alignment => alignment.Target.Name,
                alignment => GetSourceField(alignment, side) is { } sourceField
                    ? GetCell(row, sourceField.ColumnIndex)
                    : null,
                StringComparer.Ordinal);
            var recordKey = keyDefinition.FieldNames.Count == 1 &&
                            string.Equals(keyDefinition.FieldNames[0], SingletonRecordKey, StringComparison.Ordinal)
                ? new LubanRecordKey(new[] { SingletonRecordKey })
                : new LubanRecordKey(keyDefinition.FieldNames.Select(fieldName =>
                    LubanKeyValueNormalizer.Normalize(
                        targetSchema.FindField(fieldName)!,
                        cells[fieldName]?.Payload.RawValue) ?? string.Empty).ToArray());
            var key = recordKey.StableValue;
            var record = new LubanRecord(
                key,
                targetSchema.Fields
                    .Where(field => !ignoredFields.Contains(field.Name))
                    .Select(field => new KeyValuePair<string, CellPayload>(
                    field.Name,
                    cells[field.Name]?.Payload ?? CellPayload.Blank)));
            records.Add(new ParsedRecord(
                row.RowNumber,
                record,
                cells,
                IsSingletonKey(keyDefinition) ? "单例记录" : recordKey.DisplayValue,
                CreateAlignedCellArray(row, schemaMerge.GetSchema(side), schemaMerge, side)));
        }

        return new ParsedDataset(records);
    }

    private static ParsedDataset NormalizeStructurallyDeletedFieldData(
        ParsedDataset @base,
        ParsedDataset side,
        SchemaMergePlan schemaMerge,
        SchemaSide schemaSide)
    {
        var deletedAlignments = schemaMerge.Fields
            .Where(alignment => alignment.Base is not null && GetSourceField(alignment, schemaSide) is null)
            .ToArray();
        if (deletedAlignments.Length == 0)
            return side;

        var records = side.Records.Select(record =>
        {
            var baseRecord = @base.ByKey.GetValueOrDefault(record.Record.Key);
            if (baseRecord is null)
                return record;
            var fields = record.Record.Fields.ToDictionary(field => field.Key, field => field.Value, StringComparer.Ordinal);
            foreach (var alignment in deletedAlignments)
            {
                if (baseRecord.Record.Fields.TryGetValue(alignment.Target.Name, out var baseValue))
                    fields[alignment.Target.Name] = baseValue;
            }
            return record with { Record = new LubanRecord(record.Record.Key, fields) };
        }).ToArray();
        return new ParsedDataset(records);
    }

    private static CellPayload[] CreateAlignedCellArray(
        OpenXmlRowSnapshot? row,
        LubanSchema sourceSchema,
        SchemaMergePlan schemaMerge,
        SchemaSide side)
    {
        var result = Enumerable.Repeat(CellPayload.Blank, schemaMerge.ColumnCount).ToArray();
        if (row is null)
            return result;

        var sourceFieldColumns = sourceSchema.Fields.Select(field => field.ColumnIndex).ToHashSet();
        var targetFieldColumns = schemaMerge.Fields.Select(field => field.Target.ColumnIndex).ToHashSet();
        foreach (var cell in row.Cells.Where(cell =>
                     cell.ColumnIndex < result.Length &&
                     !sourceFieldColumns.Contains(cell.ColumnIndex) &&
                     !targetFieldColumns.Contains(cell.ColumnIndex)))
        {
            result[cell.ColumnIndex] = cell.Payload;
        }

        foreach (var alignment in schemaMerge.Fields)
        {
            var sourceField = GetSourceField(alignment, side);
            if (sourceField is not null)
            {
                result[alignment.Target.ColumnIndex] =
                    GetCell(row, sourceField.ColumnIndex)?.Payload ?? CellPayload.Blank;
            }
        }
        return result;
    }

    private static LubanField? GetSourceField(AlignedField alignment, SchemaSide side) => side switch
    {
        SchemaSide.Base => alignment.Base,
        SchemaSide.Local => alignment.Local,
        SchemaSide.Remote => alignment.Remote,
        _ => throw new ArgumentOutOfRangeException(nameof(side))
    };

    private static string MapFieldName(SchemaMergePlan schemaMerge, string sourceName)
    {
        var alignment = schemaMerge.Fields.FirstOrDefault(field =>
            string.Equals(field.Base?.Name, sourceName, StringComparison.Ordinal) ||
            string.Equals(field.Local?.Name, sourceName, StringComparison.Ordinal) ||
            string.Equals(field.Remote?.Name, sourceName, StringComparison.Ordinal));
        return alignment?.Target.Name ?? sourceName;
    }

    private static IReadOnlyList<WorkbookEdit> CreateFinalEdits(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        SchemaMergePlan schemaMerge,
        IReadOnlyList<WorkbookEdit> automaticEdits,
        IReadOnlyList<ResolvableMergeConflict> conflicts)
    {
        var resolvedFields = ResolveFinalFields(schemaMerge, sheets.Local);
        var structuralEdits = CreateColumnStructureEdits(sheets, schemaMerge, resolvedFields);
        var selectedEdits = conflicts
            .Where(conflict => !schemaMerge.Conflicts.Contains(conflict))
            .SelectMany(conflict => conflict.GetSelectedEdits());
        var analysisByColumn = schemaMerge.Fields.ToDictionary(field => field.Target.ColumnIndex);
        var finalByIdentity = resolvedFields.ToDictionary(
            field => field.Alignment.Identity,
            StringComparer.Ordinal);
        var remappedEdits = automaticEdits
            .Concat(selectedEdits)
            .SelectMany(edit => RemapEdit(
                edit,
                schemaMerge,
                analysisByColumn,
                finalByIdentity));
        var allEdits = remappedEdits
            .OfType<ReplaceMetadataRowsEdit>()
            .Cast<WorkbookEdit>()
            .Concat(structuralEdits)
            .Concat(remappedEdits.Where(edit => edit is not ReplaceMetadataRowsEdit))
            .ToArray();
        var metadataReplacement = allEdits.OfType<ReplaceMetadataRowsEdit>().FirstOrDefault();
        if (metadataReplacement is null)
            return allEdits;

        var delta = metadataReplacement.Rows.Count - metadataReplacement.ExistingRowCount;
        return allEdits
            .Select(edit => edit is ReplaceMetadataRowsEdit
                ? edit
                : ShiftEditRows(edit, metadataReplacement.StartRowNumber + metadataReplacement.ExistingRowCount, delta))
            .ToArray();
    }

    private static WorkbookEdit ShiftEditRows(WorkbookEdit edit, int firstMovedRow, int delta) => edit switch
    {
        SetCellEdit setCell when delta != 0 => ShiftSetCell(setCell, firstMovedRow, delta),
        DeleteRowEdit deleteRow when deleteRow.RowNumber >= firstMovedRow => deleteRow with { RowNumber = deleteRow.RowNumber + delta },
        AppendRowEdit appendRow when appendRow.SourceRowNumber is int source && source >= firstMovedRow => appendRow with { SourceRowNumber = source + delta },
        _ => edit
    };

    private static SetCellEdit ShiftSetCell(SetCellEdit edit, int firstMovedRow, int delta)
    {
        var (rowNumber, columnIndex) = CellReference.Parse(edit.Address);
        return rowNumber >= firstMovedRow
            ? edit with { Address = CellReference.Create(rowNumber + delta, columnIndex) }
            : edit;
    }

    private static IReadOnlyList<ResolvedField> ResolveFinalFields(
        SchemaMergePlan schemaMerge,
        SheetSnapshot localSheet)
    {
        var candidates = schemaMerge.Fields
            .Select(alignment => (Alignment: alignment, Resolution: alignment.ResolveStructure()))
            .Where(item => item.Resolution.Field is not null)
            .Select(item => new ResolvedField(
                item.Alignment,
                item.Resolution.Field!,
                item.Resolution.Source))
            .ToArray();
        var duplicateNames = candidates
            .GroupBy(candidate => candidate.Field.Name, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToArray();
        if (duplicateNames.Length > 0)
        {
            throw new InvalidOperationException(
                "列结构选择后存在重复字段名：" + string.Join("、", duplicateNames) + "。");
        }

        var localFieldColumns = schemaMerge.Schemas.Local.Fields
            .Select(field => field.ColumnIndex)
            .ToHashSet();
        var blockedColumns = localSheet.Rows
            .SelectMany(row => row.Cells)
            .Where(HasCellContent)
            .Select(cell => cell.ColumnIndex)
            .Where(column => !localFieldColumns.Contains(column))
            .ToHashSet();
        var usedColumns = new HashSet<int>(blockedColumns);
        var nextColumn = candidates.Select(candidate => candidate.Field.ColumnIndex)
            .Concat(blockedColumns)
            .DefaultIfEmpty(-1)
            .Max() + 1;
        var resolved = new List<ResolvedField>(candidates.Length);
        foreach (var candidate in candidates
                     .OrderBy(candidate => ChoicePriority(candidate.Source))
                     .ThenBy(candidate => candidate.Field.ColumnIndex))
        {
            var columnIndex = candidate.Field.ColumnIndex;
            if (usedColumns.Contains(columnIndex))
            {
                while (usedColumns.Contains(nextColumn))
                    nextColumn++;
                columnIndex = nextColumn++;
            }
            usedColumns.Add(columnIndex);
            resolved.Add(candidate with { Field = candidate.Field with { ColumnIndex = columnIndex } });
        }
        return resolved.OrderBy(field => field.Field.ColumnIndex).ToArray();
    }

    private static int ChoicePriority(MergeChoice choice) => choice switch
    {
        MergeChoice.Local => 0,
        MergeChoice.Base => 1,
        MergeChoice.Remote => 2,
        _ => 3
    };

    private static IReadOnlyList<WorkbookEdit> CreateColumnStructureEdits(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        SchemaMergePlan schemaMerge,
        IReadOnlyList<ResolvedField> resolvedFields)
    {
        var finalByColumn = resolvedFields.ToDictionary(field => field.Field.ColumnIndex);
        var affectedColumns = schemaMerge.Schemas.Local.Fields
            .Select(field => field.ColumnIndex)
            .Concat(resolvedFields.Select(field => field.Field.ColumnIndex))
            .Distinct()
            .Order()
            .ToArray();
        var edits = new List<WorkbookEdit>();
        foreach (var localRow in sheets.Local.Rows)
        {
            foreach (var columnIndex in affectedColumns)
            {
                var actual = GetCell(localRow, columnIndex);
                var desired = finalByColumn.TryGetValue(columnIndex, out var resolvedField)
                    ? GetStructuralSourceCell(localRow.RowNumber, resolvedField, sheets, schemaMerge)
                    : null;
                var desiredPayload = desired?.Payload ?? CellPayload.Blank;
                var desiredStyle = desired?.StyleIndex ?? actual?.StyleIndex;
                if (actual is null && desiredPayload.Kind == CellValueKind.Blank && desiredStyle is null)
                    continue;
                if (actual is not null &&
                    actual.Payload.ContentEquals(desiredPayload) &&
                    string.Equals(actual.StyleIndex, desiredStyle, StringComparison.Ordinal))
                {
                    continue;
                }
                edits.Add(new SetCellEdit(
                    sheets.Local.Name,
                    CellReference.Create(localRow.RowNumber, columnIndex),
                    desiredPayload,
                    desiredStyle));
            }
        }
        return edits;
    }

    private static OpenXmlCellSnapshot? GetStructuralSourceCell(
        int rowNumber,
        ResolvedField resolvedField,
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        SchemaMergePlan schemaMerge)
    {
        if (rowNumber >= schemaMerge.TargetSchema.DataStartRowNumber)
        {
            return resolvedField.Alignment.Local is { } localField
                ? FindRow(sheets.Local, rowNumber)?.GetCell(localField.ColumnIndex)
                : null;
        }

        var sourceField = GetSourceField(resolvedField.Alignment, resolvedField.Source);
        var sourceSheet = resolvedField.Source switch
        {
            MergeChoice.Base => sheets.Base,
            MergeChoice.Local => sheets.Local,
            MergeChoice.Remote => sheets.Remote,
            _ => throw new ArgumentOutOfRangeException()
        };
        return sourceField is null
            ? null
            : FindRow(sourceSheet, rowNumber)?.GetCell(sourceField.ColumnIndex);
    }

    private static LubanField? GetSourceField(AlignedField alignment, MergeChoice choice) => choice switch
    {
        MergeChoice.Base => alignment.Base,
        MergeChoice.Local => alignment.Local,
        MergeChoice.Remote => alignment.Remote,
        _ => throw new ArgumentOutOfRangeException(nameof(choice))
    };

    private static IEnumerable<WorkbookEdit> RemapEdit(
        WorkbookEdit edit,
        SchemaMergePlan schemaMerge,
        IReadOnlyDictionary<int, AlignedField> analysisByColumn,
        IReadOnlyDictionary<string, ResolvedField> finalByIdentity)
    {
        switch (edit)
        {
            case SetCellEdit setCell:
                {
                    var (rowNumber, columnIndex) = CellReference.Parse(setCell.Address);
                    if (!analysisByColumn.TryGetValue(columnIndex, out var alignment))
                    {
                        yield return setCell;
                        yield break;
                    }
                    if (rowNumber == schemaMerge.TargetSchema.PrimaryVariableRowNumber ||
                        rowNumber == schemaMerge.TargetSchema.TypeRowNumber)
                    {
                        yield break;
                    }
                    if (!finalByIdentity.TryGetValue(alignment.Identity, out var resolvedField))
                        yield break;
                    yield return setCell with
                    {
                        Address = CellReference.Create(rowNumber, resolvedField.Field.ColumnIndex)
                    };
                    yield break;
                }
            case AppendRowEdit appendRow:
                {
                    var cells = new List<CellWrite>();
                    foreach (var cell in appendRow.Cells)
                    {
                        if (!analysisByColumn.TryGetValue(cell.ColumnIndex, out var alignment))
                        {
                            cells.Add(cell);
                            continue;
                        }
                        if (finalByIdentity.TryGetValue(alignment.Identity, out var resolvedField))
                            cells.Add(cell with { ColumnIndex = resolvedField.Field.ColumnIndex });
                    }
                    yield return appendRow with { Cells = cells };
                    yield break;
                }
            default:
                yield return edit;
                yield break;
        }
    }

    private static IReadOnlyList<OpenXmlRowSnapshot> GetDataRows(SheetSnapshot sheet, LubanSchema schema) =>
        sheet.Rows
            .Where(row => row.RowNumber >= schema.DataStartRowNumber)
            .Where(row => !IsCommentedDataRow(row))
            .Where(row => schema.Fields.Any(field =>
            {
                var payload = GetCell(row, field.ColumnIndex)?.Payload;
                return payload is not null && payload.Kind != CellValueKind.Blank;
            }))
            .OrderBy(row => row.RowNumber)
            .ToArray();

    private static bool IsCommentedDataRow(OpenXmlRowSnapshot row)
    {
        var marker = GetCell(row, 0)?.Payload.RawValue;
        return marker?.StartsWith("##", StringComparison.Ordinal) == true;
    }

    private static OpenXmlCellSnapshot? GetCell(OpenXmlRowSnapshot row, int columnIndex) =>
        row.GetCell(columnIndex);

    private static bool HasCellContent(OpenXmlCellSnapshot cell) =>
        cell.Payload.Kind != CellValueKind.Blank;

    private static bool IsSingletonKey(RecordKeyDefinition keyDefinition) =>
        keyDefinition.FieldNames.Count == 1 &&
        string.Equals(keyDefinition.FieldNames[0], SingletonRecordKey, StringComparison.Ordinal);

    private static MergeComparison CreateComparison(
        (SheetSnapshot Base, SheetSnapshot Local, SheetSnapshot Remote) sheets,
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas,
        SchemaMergePlan schemaMerge,
        ParsedDataset @base,
        ParsedDataset local,
        ParsedDataset remote,
        IReadOnlyList<ResolvableMergeConflict> conflicts,
        IReadOnlySet<string> ignoredFields)
    {
        var schema = schemaMerge.TargetSchema;
        var columnCount = schemaMerge.ColumnCount;
        var columnHeaders = Enumerable.Range(0, columnCount).Select(MergeComparison.ColumnName).ToArray();
        var rows = new List<ComparisonRowPlan>();
        var metadataRegionConflict = conflicts.FirstOrDefault(conflict =>
            conflict.Conflict.Kind == MergeConflictKind.MetadataChanged &&
            string.Equals(conflict.Conflict.RecordKey, "Luban 元数据行", StringComparison.Ordinal));
        var metadataConflictsByRow = conflicts
            .Where(conflict => (conflict.Conflict.Kind == MergeConflictKind.MetadataChanged ||
                                schemaMerge.Conflicts.Contains(conflict)) &&
                               !ReferenceEquals(conflict, metadataRegionConflict) &&
                               conflict.RowNumber is not null)
            .GroupBy(conflict => conflict.RowNumber!.Value)
            .ToDictionary(group => group.Key, group => (IReadOnlyList<ResolvableMergeConflict>)group.ToArray());
        var dataConflictsByRecord = conflicts
            .Where(conflict => conflict.Conflict.Kind != MergeConflictKind.MetadataChanged &&
                               !schemaMerge.Conflicts.Contains(conflict))
            .GroupBy(conflict => conflict.Conflict.RecordKey, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => (IReadOnlyList<ResolvableMergeConflict>)group.ToArray(),
                StringComparer.Ordinal);

        var metadataRowCount = new[]
        {
            schemas.Base.MetadataRows.Count,
            schemas.Local.MetadataRows.Count,
            schemas.Remote.MetadataRows.Count
        }.Max();
        for (var rowNumber = 1; rowNumber <= metadataRowCount; rowNumber++)
        {
            var metadataCellConflicts = new Dictionary<int, string>();
            var rowIndex = rows.Count;
            var baseCells = CreateAlignedCellArray(
                FindMetadataRow(sheets.Base, schemas.Base, rowNumber), schemas.Base, schemaMerge, SchemaSide.Base);
            var localCells = CreateAlignedCellArray(
                FindMetadataRow(sheets.Local, schemas.Local, rowNumber), schemas.Local, schemaMerge, SchemaSide.Local);
            var remoteCells = CreateAlignedCellArray(
                FindMetadataRow(sheets.Remote, schemas.Remote, rowNumber), schemas.Remote, schemaMerge, SchemaSide.Remote);
            if (metadataRegionConflict is not null)
            {
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    if (baseCells[columnIndex].ContentEquals(localCells[columnIndex]) &&
                        baseCells[columnIndex].ContentEquals(remoteCells[columnIndex]))
                    {
                        continue;
                    }
                    metadataCellConflicts[columnIndex] = metadataRegionConflict.Id;
                    if (metadataRegionConflict.GridRowIndex < 0)
                        metadataRegionConflict.SetGridLocation(rowIndex, columnIndex);
                }
            }
            foreach (var conflict in metadataConflictsByRow.GetValueOrDefault(rowNumber) ??
                                     Array.Empty<ResolvableMergeConflict>())
            {
                var address = conflict.Conflict.FieldName!;
                var (_, columnIndex) = CellReference.Parse(address);
                metadataCellConflicts[columnIndex] = conflict.Id;
                conflict.SetGridLocation(rowIndex, columnIndex);
            }
            rows.Add(new ComparisonRowPlan(
                true,
                $"structure:{rowNumber}",
                rowNumber,
                rowNumber,
                rowNumber,
                baseCells,
                localCells,
                remoteCells,
                null,
                metadataCellConflicts));
        }

        var consumedLocalRows = new HashSet<int>();
        var consumedRemoteRows = new HashSet<int>();
        foreach (var baseRecord in @base.Records.OrderBy(record => record.RowNumber))
        {
            var localRecord = local.ByKey.GetValueOrDefault(baseRecord.Record.Key);
            var remoteRecord = remote.ByKey.GetValueOrDefault(baseRecord.Record.Key);
            var matchingConflicts = dataConflictsByRecord.GetValueOrDefault(baseRecord.DisplayKey) ??
                                    Array.Empty<ResolvableMergeConflict>();
            var isDivergentKeyConflict = localRecord is null && remoteRecord is null &&
                                         matchingConflicts.Any(conflict => conflict.Conflict.FieldName is not null);
            if (isDivergentKeyConflict)
            {
                localRecord = local.ByRow.GetValueOrDefault(baseRecord.RowNumber);
                remoteRecord = remote.ByRow.GetValueOrDefault(baseRecord.RowNumber);
                if (localRecord is not null && consumedLocalRows.Contains(localRecord.RowNumber))
                    localRecord = null;
                if (remoteRecord is not null && consumedRemoteRows.Contains(remoteRecord.RowNumber))
                    remoteRecord = null;
            }

            if (localRecord is not null)
                consumedLocalRows.Add(localRecord.RowNumber);
            if (remoteRecord is not null)
                consumedRemoteRows.Add(remoteRecord.RowNumber);
            AddComparisonRow(rows, baseRecord, localRecord, remoteRecord, schema, columnCount, dataConflictsByRecord);
        }

        foreach (var localRecord in local.Records
                     .Where(record => !consumedLocalRows.Contains(record.RowNumber))
                     .OrderBy(record => record.RowNumber))
        {
            var remoteRecord = remote.ByKey.GetValueOrDefault(localRecord.Record.Key);
            if (remoteRecord is not null && consumedRemoteRows.Contains(remoteRecord.RowNumber))
                remoteRecord = null;
            consumedLocalRows.Add(localRecord.RowNumber);
            if (remoteRecord is not null)
                consumedRemoteRows.Add(remoteRecord.RowNumber);
            AddComparisonRow(rows, null, localRecord, remoteRecord, schema, columnCount, dataConflictsByRecord);
        }

        foreach (var remoteRecord in remote.Records
                     .Where(record => !consumedRemoteRows.Contains(record.RowNumber))
                     .OrderBy(record => record.RowNumber))
        {
            consumedRemoteRows.Add(remoteRecord.RowNumber);
            AddComparisonRow(rows, null, null, remoteRecord, schema, columnCount, dataConflictsByRecord);
        }

        var ignoredColumns = schema.Fields
            .Where(field => ignoredFields.Contains(field.Name))
            .Select(field => field.ColumnIndex)
            .ToHashSet();
        return new MergeComparison(
            columnHeaders,
            rows,
            conflicts,
            ignoredColumns,
            schemaMerge.IsIncludedForPreview,
            schemaMerge.MergeStructuralCellForPreview);
    }

    private static void AddComparisonRow(
        ICollection<ComparisonRowPlan> rows,
        ParsedRecord? baseRecord,
        ParsedRecord? localRecord,
        ParsedRecord? remoteRecord,
        LubanSchema schema,
        int columnCount,
        IReadOnlyDictionary<string, IReadOnlyList<ResolvableMergeConflict>> conflictsByRecord)
    {
        var recordKey = baseRecord?.DisplayKey ?? localRecord?.DisplayKey ?? remoteRecord?.DisplayKey ?? string.Empty;
        var matchingConflicts = conflictsByRecord.GetValueOrDefault(recordKey) ??
                                Array.Empty<ResolvableMergeConflict>();
        var divergentKeyConflict = baseRecord is not null && localRecord is not null && remoteRecord is not null &&
                                   baseRecord.Record.Key != localRecord.Record.Key &&
                                   baseRecord.Record.Key != remoteRecord.Record.Key;
        var rowConflict = matchingConflicts.FirstOrDefault(conflict => conflict.Conflict.FieldName is null) ??
                          (divergentKeyConflict ? matchingConflicts.FirstOrDefault() : null);
        var cellConflicts = new Dictionary<int, string>();
        var rowIndex = rows.Count;
        foreach (var conflict in matchingConflicts.Where(conflict =>
                     conflict.Conflict.FieldName is not null && !ReferenceEquals(conflict, rowConflict)))
        {
            var columnIndex = schema.FindField(conflict.Conflict.FieldName!)?.ColumnIndex ?? 0;
            cellConflicts[columnIndex] = conflict.Id;
            conflict.SetGridLocation(rowIndex, columnIndex);
        }
        if (rowConflict is not null)
        {
            var conflictColumn = rowConflict.Conflict.FieldName is { } fieldName
                ? schema.FindField(fieldName)?.ColumnIndex ?? 0
                : schema.Fields.FirstOrDefault()?.ColumnIndex ?? 0;
            rowConflict.SetGridLocation(rowIndex, conflictColumn);
        }

        rows.Add(new ComparisonRowPlan(
            false,
            recordKey,
            baseRecord?.RowNumber,
            localRecord?.RowNumber,
            remoteRecord?.RowNumber,
            baseRecord?.AlignedCells,
            localRecord?.AlignedCells,
            remoteRecord?.AlignedCells,
            rowConflict?.Id,
            cellConflicts));
    }

    private static OpenXmlRowSnapshot? FindRow(SheetSnapshot sheet, int rowNumber) => sheet.GetRow(rowNumber);

    private static OpenXmlRowSnapshot? FindMetadataRow(
        SheetSnapshot sheet,
        LubanSchema schema,
        int rowNumber) =>
        schema.MetadataRows.Any(row => row.RowNumber == rowNumber)
            ? sheet.GetRow(rowNumber)
            : null;

    private static MergeEditPlan CreateEdits(
        ParsedDataset @base,
        ParsedDataset local,
        ParsedDataset remote,
        LubanSchema localSchema,
        RecordKeyDefinition keyDefinition,
        string sheetName,
        IReadOnlySet<string> ignoredFields)
    {
        var edits = new List<WorkbookEdit>();
        var conflicts = new List<ResolvableMergeConflict>();
        var changedCells = 0;
        var addedRecords = 0;
        var deletedRecords = 0;
        var divergentKeyPlans = DetectDivergentKeyChanges(
            @base, local, remote, keyDefinition, localSchema, sheetName, ignoredFields);
        conflicts.AddRange(divergentKeyPlans.Conflicts);
        var keys = local.Records.Select(record => record.Record.Key)
            .Concat(@base.Records.Select(record => record.Record.Key))
            .Concat(remote.Records.Select(record => record.Record.Key))
            .Distinct(StringComparer.Ordinal)
            .Where(key => !divergentKeyPlans.ConsumedKeys.Contains(key));

        foreach (var key in keys)
        {
            var baseRecord = @base.ByKey.GetValueOrDefault(key);
            var localRecord = local.ByKey.GetValueOrDefault(key);
            var remoteRecord = remote.ByKey.GetValueOrDefault(key);

            if (baseRecord is null)
            {
                if (localRecord is null && remoteRecord is not null)
                {
                    edits.Add(CreateAppendEdit(remoteRecord, localSchema, sheetName));
                    addedRecords++;
                }
                else if (localRecord is not null && remoteRecord is not null &&
                         !localRecord.Record.ContentEquals(remoteRecord.Record))
                {
                    var conflict = new MergeConflict(
                        MergeConflictKind.AddAdd,
                        localRecord.DisplayKey,
                        null,
                        $"LOCAL 和 REMOTE 新增了同键但内容不同的记录 {localRecord.DisplayKey}。");
                    conflicts.Add(CreateResolution(
                        conflicts.Count,
                        conflict,
                        localRecord.RowNumber,
                        "<不存在>",
                        FormatRecord(localRecord.Record),
                        FormatRecord(remoteRecord.Record),
                        new DeleteRowEdit(sheetName, localRecord.RowNumber),
                        Array.Empty<WorkbookEdit>(),
                        CreateSetRecordEdits(localRecord, remoteRecord.Record, localSchema, sheetName)));
                }
                continue;
            }

            if (localRecord is null && remoteRecord is null)
                continue;

            if (localRecord is null)
            {
                if (!remoteRecord!.Record.ContentEquals(baseRecord.Record))
                {
                    var conflict = DeleteModifyConflict(baseRecord.DisplayKey);
                    conflicts.Add(CreateResolution(
                        conflicts.Count,
                        conflict,
                        baseRecord.RowNumber,
                        FormatRecord(baseRecord.Record),
                        "<已删除>",
                        FormatRecord(remoteRecord.Record),
                        CreateAppendEdit(baseRecord, localSchema, sheetName),
                        Array.Empty<WorkbookEdit>(),
                        new WorkbookEdit[] { CreateAppendEdit(remoteRecord, localSchema, sheetName) }));
                }
                continue;
            }

            if (remoteRecord is null)
            {
                if (localRecord.Record.ContentEquals(baseRecord.Record))
                {
                    edits.Add(new DeleteRowEdit(sheetName, localRecord.RowNumber));
                    deletedRecords++;
                }
                else
                {
                    var conflict = DeleteModifyConflict(baseRecord.DisplayKey);
                    conflicts.Add(CreateResolution(
                        conflicts.Count,
                        conflict,
                        localRecord.RowNumber,
                        FormatRecord(baseRecord.Record),
                        FormatRecord(localRecord.Record),
                        "<已删除>",
                        CreateSetRecordEdits(localRecord, baseRecord.Record, localSchema, sheetName),
                        Array.Empty<WorkbookEdit>(),
                        new DeleteRowEdit(sheetName, localRecord.RowNumber)));
                }
                continue;
            }

            foreach (var field in localSchema.Fields.Where(field => !ignoredFields.Contains(field.Name)))
            {
                var baseCell = baseRecord.Record.Fields[field.Name];
                var localCell = localRecord.Record.Fields[field.Name];
                var remoteCell = remoteRecord.Record.Fields[field.Name];
                var decision = CellThreeWayMerger.Merge(baseCell, localCell, remoteCell, baseRecord.DisplayKey, field.Name);
                var address = CellReference.Create(localRecord.RowNumber, field.ColumnIndex);
                if (decision.Conflict is not null)
                {
                    conflicts.Add(CreateResolution(
                        conflicts.Count,
                        decision.Conflict,
                        localRecord.RowNumber,
                        FormatCell(baseCell),
                        FormatCell(localCell),
                        FormatCell(remoteCell),
                        CreateSetCellChoice(sheetName, address, localCell, baseCell),
                        Array.Empty<WorkbookEdit>(),
                        CreateSetCellChoice(sheetName, address, localCell, remoteCell)));
                }
                else if (!decision.Result!.ContentEquals(localCell))
                {
                    var styleIndex = localRecord.Cells[field.Name] is null &&
                                     decision.Result.ContentEquals(remoteCell)
                        ? remoteRecord.Cells[field.Name]?.StyleIndex
                        : null;
                    edits.Add(new SetCellEdit(sheetName, address, decision.Result, styleIndex));
                    changedCells++;
                }
            }
        }

        return new MergeEditPlan(edits, conflicts, changedCells, addedRecords, deletedRecords);
    }

    private static DivergentKeyPlan DetectDivergentKeyChanges(
        ParsedDataset @base,
        ParsedDataset local,
        ParsedDataset remote,
        RecordKeyDefinition keyDefinition,
        LubanSchema localSchema,
        string sheetName,
        IReadOnlySet<string> ignoredFields)
    {
        var conflicts = new List<ResolvableMergeConflict>();
        var consumedKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var baseRecord in @base.Records)
        {
            if (local.ByKey.ContainsKey(baseRecord.Record.Key) || remote.ByKey.ContainsKey(baseRecord.Record.Key))
                continue;

            var localAtSourceRow = local.ByRow.GetValueOrDefault(baseRecord.RowNumber);
            var remoteAtSourceRow = remote.ByRow.GetValueOrDefault(baseRecord.RowNumber);
            if (localAtSourceRow is null || remoteAtSourceRow is null ||
                localAtSourceRow.Record.Key == remoteAtSourceRow.Record.Key)
                continue;

            var differingField = keyDefinition.FieldNames.FirstOrDefault(field =>
                localAtSourceRow.Cells[field]?.Payload.RawValue != remoteAtSourceRow.Cells[field]?.Payload.RawValue);
            if (differingField is null)
                continue;

            var displayKey = string.Join(" | ", keyDefinition.FieldNames.Select(field =>
                baseRecord.Cells[field]?.Payload.RawValue ?? string.Empty));
            var conflict = new MergeConflict(
                MergeConflictKind.CellChangedDifferently,
                displayKey,
                differingField,
                $"记录 {displayKey} 的主键字段 {differingField} 在 LOCAL 和 REMOTE 中被修改为不同内容。");
            conflicts.Add(CreateResolution(
                conflicts.Count,
                conflict,
                localAtSourceRow.RowNumber,
                FormatRecord(baseRecord.Record),
                FormatRecord(localAtSourceRow.Record),
                FormatRecord(remoteAtSourceRow.Record),
                CreateSetRecordEdits(localAtSourceRow, baseRecord.Record, localSchema, sheetName, ignoredFields),
                Array.Empty<WorkbookEdit>(),
                CreateSetRecordEdits(localAtSourceRow, remoteAtSourceRow.Record, localSchema, sheetName, ignoredFields)));
            consumedKeys.Add(baseRecord.Record.Key);
            consumedKeys.Add(localAtSourceRow.Record.Key);
            consumedKeys.Add(remoteAtSourceRow.Record.Key);
        }

        return new DivergentKeyPlan(conflicts, consumedKeys);
    }

    private static MergeConflict DeleteModifyConflict(string key) => new(
        MergeConflictKind.DeleteModify,
        key,
        null,
        $"记录 {key} 在一侧被删除、另一侧被修改。");

    private static ResolvableMergeConflict CreateResolution(
        int index,
        MergeConflict conflict,
        int? rowNumber,
        string baseValue,
        string localValue,
        string remoteValue,
        WorkbookEdit baseEdit,
        IReadOnlyList<WorkbookEdit> localEdits,
        IReadOnlyList<WorkbookEdit> remoteEdits) =>
        CreateResolution(index, conflict, rowNumber, baseValue, localValue, remoteValue,
            new[] { baseEdit }, localEdits, remoteEdits);

    private static ResolvableMergeConflict CreateResolution(
        int index,
        MergeConflict conflict,
        int? rowNumber,
        string baseValue,
        string localValue,
        string remoteValue,
        IReadOnlyList<WorkbookEdit> baseEdits,
        IReadOnlyList<WorkbookEdit> localEdits,
        WorkbookEdit remoteEdit) =>
        CreateResolution(index, conflict, rowNumber, baseValue, localValue, remoteValue,
            baseEdits, localEdits, new[] { remoteEdit });

    private static ResolvableMergeConflict CreateResolution(
        int index,
        MergeConflict conflict,
        int? rowNumber,
        string baseValue,
        string localValue,
        string remoteValue,
        IReadOnlyList<WorkbookEdit> baseEdits,
        IReadOnlyList<WorkbookEdit> localEdits,
        IReadOnlyList<WorkbookEdit> remoteEdits) => new(
            $"{conflict.Kind}:{conflict.RecordKey}:{conflict.FieldName ?? "record"}:{index}",
            conflict,
            rowNumber,
            baseValue,
            localValue,
            remoteValue,
            new Dictionary<MergeChoice, IReadOnlyList<WorkbookEdit>>
            {
                [MergeChoice.Base] = baseEdits,
                [MergeChoice.Local] = localEdits,
                [MergeChoice.Remote] = remoteEdits
            });

    private static IReadOnlyList<WorkbookEdit> CreateSetCellChoice(
        string sheetName,
        string address,
        CellPayload local,
        CellPayload target) =>
        local.ContentEquals(target)
            ? Array.Empty<WorkbookEdit>()
            : new WorkbookEdit[] { new SetCellEdit(sheetName, address, target) };

    private static AppendRowEdit CreateAppendEdit(ParsedRecord source, LubanSchema schema, string sheetName) =>
        new(sheetName, schema.Fields
            .Where(field => source.Record.Fields[field.Name].Kind != CellValueKind.Blank)
            .Select(field => new CellWrite(
                field.ColumnIndex,
                source.Record.Fields[field.Name],
                source.Cells[field.Name]?.StyleIndex))
            .ToArray(),
            source.RowNumber);

    private static IReadOnlyList<WorkbookEdit> CreateSetRecordEdits(
        ParsedRecord local,
        LubanRecord target,
        LubanSchema schema,
        string sheetName,
        IReadOnlySet<string>? ignoredFields = null) => schema.Fields
            .Where(field => ignoredFields is null || !ignoredFields.Contains(field.Name))
            .Where(field => !local.Record.Fields[field.Name].ContentEquals(target.Fields[field.Name]))
            .Select(field => (WorkbookEdit)new SetCellEdit(
                sheetName,
                CellReference.Create(local.RowNumber, field.ColumnIndex),
                target.Fields[field.Name]))
            .ToArray();

    private static string FormatCell(CellPayload cell) => cell.Kind switch
    {
        CellValueKind.Blank => "<空>",
        CellValueKind.Formula => $"={cell.FormulaText}  [缓存: {cell.CachedValue}]",
        _ => cell.RawValue ?? "<空>"
    };

    private static string FormatRecord(LubanRecord record) => string.Join(" | ",
        record.Fields.Select(field => $"{field.Key}={FormatCell(field.Value)}"));

    private static string FormatKeyFailures(IReadOnlyList<RecordKeyValidationResult> attempts)
    {
        var issues = attempts.SelectMany(attempt => attempt.Issues).Take(20).Select(issue => issue.Message);
        return "没有索引能在 BASE、LOCAL、REMOTE 中同时满足唯一性：" + string.Join("；", issues);
    }

    private sealed record ParsedRecord(
        int RowNumber,
        LubanRecord Record,
        IReadOnlyDictionary<string, OpenXmlCellSnapshot?> Cells,
        string DisplayKey,
        CellPayload[] AlignedCells);

    private enum SchemaSide
    {
        Base,
        Local,
        Remote
    }

    private sealed record FieldAlignmentDraft(
        string Identity,
        string DisplayName,
        LubanField? Base,
        LubanField? Local,
        LubanField? Remote);

    private sealed class MutableFieldAlignment
    {
        public MutableFieldAlignment(
            string identity,
            string displayName,
            LubanField? @base,
            LubanField? local,
            LubanField? remote)
        {
            Identity = identity;
            DisplayName = displayName;
            Base = @base;
            Local = local;
            Remote = remote;
        }

        public string Identity { get; }
        public string DisplayName { get; }
        public LubanField? Base { get; }
        public LubanField? Local { get; set; }
        public LubanField? Remote { get; set; }
        public FieldAlignmentDraft ToImmutable() =>
            new(Identity, DisplayName, Base, Local, Remote);
    }

    private sealed record FieldStructureDecision(
        LubanField? Result,
        MergeChoice Source,
        bool IsConflict);

    private sealed record ResolvedField(
        AlignedField Alignment,
        LubanField Field,
        MergeChoice Source);

    private sealed record AlignedField(
        string Identity,
        LubanField Target,
        LubanField? Base,
        LubanField? Local,
        LubanField? Remote,
        FieldStructureDecision StructureDecision,
        ResolvableMergeConflict? StructureConflict)
    {
        public bool HasStructuralVariation =>
            !FieldEquals(Base, Local) || !FieldEquals(Base, Remote);

        public (LubanField? Field, MergeChoice Source) ResolveStructure()
        {
            if (StructureConflict?.SelectedChoice is { } selected)
            {
                return selected switch
                {
                    MergeChoice.Base => (Base, selected),
                    MergeChoice.Local => (Local, selected),
                    MergeChoice.Remote => (Remote, selected),
                    _ => throw new ArgumentOutOfRangeException()
                };
            }
            return (StructureDecision.Result, StructureDecision.Source);
        }
    }

    private sealed record SchemaMergePlan(
        LubanSchema TargetSchema,
        IReadOnlyList<AlignedField> Fields,
        int ColumnCount,
        IReadOnlyList<ResolvableMergeConflict> Conflicts,
        IReadOnlyList<string> StructuralChanges,
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) Schemas)
    {
        public LubanSchema GetSchema(SchemaSide side) => side switch
        {
            SchemaSide.Base => Schemas.Base,
            SchemaSide.Local => Schemas.Local,
            SchemaSide.Remote => Schemas.Remote,
            _ => throw new ArgumentOutOfRangeException(nameof(side))
        };

        public bool IsIncludedForPreview(int columnIndex)
        {
            var alignment = Fields.FirstOrDefault(field => field.Target.ColumnIndex == columnIndex);
            if (alignment is null)
                return true;
            if (alignment.StructureConflict is null)
                return alignment.StructureDecision.Result is not null;
            return (alignment.StructureConflict.SelectedChoice ?? MergeChoice.Local) switch
            {
                MergeChoice.Base => alignment.Base is not null,
                MergeChoice.Local => alignment.Local is not null,
                MergeChoice.Remote => alignment.Remote is not null,
                _ => true
            };
        }

        public CellPayload? MergeStructuralCellForPreview(
            int columnIndex,
            string recordKey,
            CellPayload baseCell,
            CellPayload localCell,
            CellPayload remoteCell)
        {
            var alignment = Fields.FirstOrDefault(field => field.Target.ColumnIndex == columnIndex);
            if (alignment?.Base is null || alignment.Local is not null && alignment.Remote is not null)
                return null;
            var effectiveLocal = alignment.Local is null ? baseCell : localCell;
            var effectiveRemote = alignment.Remote is null ? baseCell : remoteCell;
            var decision = CellThreeWayMerger.Merge(
                baseCell,
                effectiveLocal,
                effectiveRemote,
                recordKey,
                alignment.Target.Name);
            return decision.Result ?? effectiveLocal;
        }
    }

    private sealed record MetadataMergePlan(
        IReadOnlyList<WorkbookEdit> Edits,
        IReadOnlyList<ResolvableMergeConflict> Conflicts);

    private sealed class ParsedDataset
    {
        public ParsedDataset(IReadOnlyList<ParsedRecord> records)
        {
            Records = records;
            ByKey = records.ToDictionary(record => record.Record.Key, StringComparer.Ordinal);
            ByRow = records.ToDictionary(record => record.RowNumber);
        }

        public IReadOnlyList<ParsedRecord> Records { get; }
        public IReadOnlyDictionary<string, ParsedRecord> ByKey { get; }
        public IReadOnlyDictionary<int, ParsedRecord> ByRow { get; }
    }

    private sealed record MergeEditPlan(
        IReadOnlyList<WorkbookEdit> Edits,
        IReadOnlyList<ResolvableMergeConflict> Conflicts,
        int ChangedCells,
        int AddedRecords,
        int DeletedRecords);

    private sealed record DivergentKeyPlan(
        IReadOnlyList<ResolvableMergeConflict> Conflicts,
        IReadOnlySet<string> ConsumedKeys);
}
