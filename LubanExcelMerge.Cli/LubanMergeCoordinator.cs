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
    MergePreparationTimings PreparationTimings)
{
    public bool Succeeded => Conflicts.Count == 0 && OutputPath is not null;
}

public sealed class LubanMergeCoordinator
{
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
            session.PreparationTimings);
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
            var schemaMerge = CreateSchemaMergePlan(schemas, sheets.Local, sheets.Remote);
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
                CreateRecordSources(sheets.Base, schemas.Base, basePath, dataRows[0]),
                CreateRecordSources(sheets.Local, schemas.Local, localPath, dataRows[1]),
                CreateRecordSources(sheets.Remote, schemas.Remote, remotePath, dataRows[2])
            };
            var keyCandidates = configuredKey is null
                ? logicalTable.DeclaredIndexes
                : new[] { new RecordKeyDefinition(configuredKey) };
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

            var keyDefinition = keySelection.Selected;
            ValidateIgnoredFields(schemaMerge.TargetSchema, keyDefinition, ignoredFields, options.LoadedConfigPath);
            if (options.ValidateLogicalTableUniqueness)
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
            var ignoredFieldSet = ignoredFields.ToHashSet(StringComparer.Ordinal);
            var datasets = new[]
            {
                CreateDataset(schemas.Base, schemaMerge, keyDefinition, ignoredFieldSet, dataRows[0]),
                CreateDataset(schemas.Local, schemaMerge, keyDefinition, ignoredFieldSet, dataRows[1]),
                CreateDataset(schemas.Remote, schemaMerge, keyDefinition, ignoredFieldSet, dataRows[2])
            };
            var plan = CreateEdits(
                datasets[0],
                datasets[1],
                datasets[2],
                schemaMerge.TargetSchema,
                keyDefinition,
                sheets.Local.Name,
                ignoredFieldSet);
            var conflicts = metadataPlan.Conflicts.Concat(plan.Conflicts).ToArray();
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
                metadataPlan.Conflicts.Count);
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
        if (!string.IsNullOrEmpty(table.Mode) && !string.Equals(table.Mode, "map", StringComparison.OrdinalIgnoreCase))
            throw new UnsafeWorkbookException($"逻辑表 {table.FullName} 的模式 {table.Mode} 尚不能安全地在无界面模式下合并。");
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
            return (ParseSchema(@base), ParseSchema(local), ParseSchema(remote));
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

    private static SchemaMergePlan CreateSchemaMergePlan(
        (LubanSchema Base, LubanSchema Local, LubanSchema Remote) schemas,
        SheetSnapshot localSheet,
        SheetSnapshot remoteSheet)
    {
        if (schemas.Base.IsRestricted || schemas.Local.IsRestricted || schemas.Remote.IsRestricted)
        {
            var restrictions = schemas.Base.Restrictions
                .Concat(schemas.Local.Restrictions)
                .Concat(schemas.Remote.Restrictions)
                .Distinct(StringComparer.Ordinal);
            throw new UnsafeWorkbookException("工作表结构处于受限模式：" + string.Join("；", restrictions));
        }

        if (schemas.Base.PrimaryVariableRowNumber != schemas.Local.PrimaryVariableRowNumber ||
            schemas.Base.PrimaryVariableRowNumber != schemas.Remote.PrimaryVariableRowNumber ||
            schemas.Base.TypeRowNumber != schemas.Local.TypeRowNumber ||
            schemas.Base.TypeRowNumber != schemas.Remote.TypeRowNumber ||
            schemas.Base.DataStartRowNumber != schemas.Local.DataStartRowNumber ||
            schemas.Base.DataStartRowNumber != schemas.Remote.DataStartRowNumber)
        {
            throw new UnsafeWorkbookException(
                "BASE、LOCAL、REMOTE 的 Luban 表头行或数据起始行不一致；" +
                "当前版本支持追加字段列，但仍不支持插入或删除元数据行。");
        }

        var baseFields = schemas.Base.Fields.ToDictionary(field => field.Name, StringComparer.Ordinal);
        var localFields = schemas.Local.Fields.ToDictionary(field => field.Name, StringComparer.Ordinal);
        var remoteFields = schemas.Remote.Fields.ToDictionary(field => field.Name, StringComparer.Ordinal);
        foreach (var baseField in schemas.Base.Fields)
        {
            ValidateExistingField(baseField, localFields.GetValueOrDefault(baseField.Name), "LOCAL");
            ValidateExistingField(baseField, remoteFields.GetValueOrDefault(baseField.Name), "REMOTE");
        }
        var baseFieldOrder = schemas.Base.Fields.Select(field => field.Name).ToArray();
        foreach (var (label, fields) in new[] { ("LOCAL", schemas.Local.Fields), ("REMOTE", schemas.Remote.Fields) })
        {
            var existingFieldOrder = fields
                .Where(field => baseFields.ContainsKey(field.Name))
                .Select(field => field.Name)
                .ToArray();
            if (!baseFieldOrder.SequenceEqual(existingFieldOrder, StringComparer.Ordinal))
            {
                throw new UnsafeWorkbookException(
                    $"{label} 调整了既有字段的相对顺序；当前版本支持新增字段列，但不自动合并既有字段重排。");
            }
        }

        var occupiedLocalColumns = localSheet.Rows
            .SelectMany(row => row.Cells)
            .Select(cell => cell.ColumnIndex)
            .ToHashSet();
        var usedTargetColumns = schemas.Local.Fields.Select(field => field.ColumnIndex).ToHashSet();
        var nextTargetColumn = occupiedLocalColumns.DefaultIfEmpty(-1).Max() + 1;
        var alignments = new List<AlignedField>();

        foreach (var localField in schemas.Local.Fields)
        {
            baseFields.TryGetValue(localField.Name, out var baseField);
            remoteFields.TryGetValue(localField.Name, out var remoteField);
            alignments.Add(new AlignedField(localField, baseField, localField, remoteField));
        }

        foreach (var remoteField in schemas.Remote.Fields.Where(field => !localFields.ContainsKey(field.Name)))
        {
            var targetColumn = remoteField.ColumnIndex;
            if (occupiedLocalColumns.Contains(targetColumn) || usedTargetColumns.Contains(targetColumn))
            {
                while (occupiedLocalColumns.Contains(nextTargetColumn) || usedTargetColumns.Contains(nextTargetColumn))
                    nextTargetColumn++;
                targetColumn = nextTargetColumn++;
            }

            if (targetColumn != remoteField.ColumnIndex && remoteSheet.Rows.Any(row =>
                    GetCell(row, remoteField.ColumnIndex)?.Payload.Kind == CellValueKind.Formula))
            {
                throw new UnsafeWorkbookException(
                    $"REMOTE 新增公式字段 {remoteField.Name} 需要从 {MergeComparison.ColumnName(remoteField.ColumnIndex)} 列" +
                    $"移动到 {MergeComparison.ColumnName(targetColumn)} 列；为避免公式引用失真，请先将该字段追加到空闲列后再合并。");
            }

            var targetField = remoteField with { ColumnIndex = targetColumn };
            usedTargetColumns.Add(targetColumn);
            alignments.Add(new AlignedField(targetField, null, null, remoteField));
        }

        alignments.Sort((left, right) => left.Target.ColumnIndex.CompareTo(right.Target.ColumnIndex));
        var targetSchema = new LubanSchema(
            alignments.Select(alignment => alignment.Target).ToArray(),
            schemas.Local.MetadataRows,
            schemas.Local.PrimaryVariableRowNumber,
            schemas.Local.TypeRowNumber,
            schemas.Local.DataStartRowNumber,
            false,
            Array.Empty<string>());
        var maximumTargetColumn = alignments.Select(alignment => alignment.Target.ColumnIndex).DefaultIfEmpty(-1).Max();
        var maximumRawColumn = localSheet.Rows
            .SelectMany(row => row.Cells)
            .Select(cell => cell.ColumnIndex)
            .DefaultIfEmpty(-1)
            .Max();
        return new SchemaMergePlan(
            targetSchema,
            alignments,
            Math.Max(maximumTargetColumn, maximumRawColumn) + 1);
    }

    private static void ValidateExistingField(LubanField expected, LubanField? actual, string side)
    {
        if (actual is null)
        {
            throw new UnsafeWorkbookException(
                $"{side} 删除或重命名了既有字段 {expected.Name}；当前版本只支持新增字段列。");
        }
        if (!string.Equals(expected.TypeName, actual.TypeName, StringComparison.Ordinal))
        {
            throw new UnsafeWorkbookException(
                $"{side} 修改了既有字段 {expected.Name} 的类型；当前版本只支持新增字段列。");
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
        var alignmentByColumn = schemaMerge.Fields.ToDictionary(
            alignment => alignment.Target.ColumnIndex);
        for (var rowNumber = 1; rowNumber < schemaMerge.TargetSchema.DataStartRowNumber; rowNumber++)
        {
            var baseCells = CreateAlignedCellArray(
                FindRow(sheets.Base, rowNumber), schemas.Base, schemaMerge);
            var localCells = CreateAlignedCellArray(
                FindRow(sheets.Local, rowNumber), schemas.Local, schemaMerge);
            var remoteCells = CreateAlignedCellArray(
                FindRow(sheets.Remote, rowNumber), schemas.Remote, schemaMerge);
            for (var columnIndex = 0; columnIndex < schemaMerge.ColumnCount; columnIndex++)
            {
                var baseCell = baseCells[columnIndex];
                var localCell = localCells[columnIndex];
                var remoteCell = remoteCells[columnIndex];
                if (baseCell.ContentEquals(localCell) && baseCell.ContentEquals(remoteCell))
                    continue;

                var address = CellReference.Create(rowNumber, columnIndex);
                var alignment = alignmentByColumn.GetValueOrDefault(columnIndex);
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

    private static ParsedDataset CreateDataset(
        LubanSchema sourceSchema,
        SchemaMergePlan schemaMerge,
        RecordKeyDefinition keyDefinition,
        IReadOnlySet<string> ignoredFields,
        IReadOnlyList<OpenXmlRowSnapshot> dataRows)
    {
        var targetSchema = schemaMerge.TargetSchema;
        var records = new List<ParsedRecord>();
        foreach (var row in dataRows)
        {
            var cells = targetSchema.Fields.ToDictionary(
                field => field.Name,
                field => sourceSchema.FindField(field.Name) is { } sourceField
                    ? GetCell(row, sourceField.ColumnIndex)
                    : null,
                StringComparer.Ordinal);
            var components = keyDefinition.FieldNames.Select(fieldName =>
                LubanKeyValueNormalizer.Normalize(
                    targetSchema.FindField(fieldName)!,
                    cells[fieldName]?.Payload.RawValue) ?? string.Empty).ToArray();
            var recordKey = new LubanRecordKey(components);
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
                recordKey.DisplayValue,
                CreateAlignedCellArray(row, sourceSchema, schemaMerge)));
        }

        return new ParsedDataset(records);
    }

    private static CellPayload[] CreateAlignedCellArray(
        OpenXmlRowSnapshot? row,
        LubanSchema sourceSchema,
        SchemaMergePlan schemaMerge)
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
            var sourceField = GetSourceField(alignment, sourceSchema);
            if (sourceField is not null)
            {
                result[alignment.Target.ColumnIndex] =
                    GetCell(row, sourceField.ColumnIndex)?.Payload ?? CellPayload.Blank;
            }
        }
        return result;
    }

    private static LubanField? GetSourceField(AlignedField alignment, LubanSchema sourceSchema)
        => sourceSchema.FindField(alignment.Target.Name);

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
        var metadataConflictsByRow = conflicts
            .Where(conflict => conflict.Conflict.Kind == MergeConflictKind.MetadataChanged &&
                               conflict.RowNumber is not null)
            .GroupBy(conflict => conflict.RowNumber!.Value)
            .ToDictionary(group => group.Key, group => (IReadOnlyList<ResolvableMergeConflict>)group.ToArray());
        var dataConflictsByRecord = conflicts
            .Where(conflict => conflict.Conflict.Kind != MergeConflictKind.MetadataChanged)
            .GroupBy(conflict => conflict.Conflict.RecordKey, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => (IReadOnlyList<ResolvableMergeConflict>)group.ToArray(),
                StringComparer.Ordinal);

        for (var rowNumber = 1; rowNumber < schema.DataStartRowNumber; rowNumber++)
        {
            var metadataCellConflicts = new Dictionary<int, string>();
            var rowIndex = rows.Count;
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
                CreateAlignedCellArray(FindRow(sheets.Base, rowNumber), schemas.Base, schemaMerge),
                CreateAlignedCellArray(FindRow(sheets.Local, rowNumber), schemas.Local, schemaMerge),
                CreateAlignedCellArray(FindRow(sheets.Remote, rowNumber), schemas.Remote, schemaMerge),
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
        return new MergeComparison(columnHeaders, rows, conflicts, ignoredColumns);
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

    private sealed record AlignedField(
        LubanField Target,
        LubanField? Base,
        LubanField? Local,
        LubanField? Remote);

    private sealed record SchemaMergePlan(
        LubanSchema TargetSchema,
        IReadOnlyList<AlignedField> Fields,
        int ColumnCount);

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
