using System.Text.Json;
using System.Text.Json.Serialization;

namespace LubanExcelMerge.Cli;

public sealed class LubanMergeConfiguration
{
    public string? DataRoot { get; init; }
    public string? TablesFile { get; init; }
    public string? FormulaRecalculation { get; init; }
    public bool? ValidateLogicalTableUniqueness { get; init; }
    public IReadOnlyList<string>? InactivePaths { get; init; }
    public IReadOnlyDictionary<string, string[]>? KeyOverrides { get; init; }
    public IReadOnlyDictionary<string, string[]>? IgnoredFields { get; init; }
    public ValidationConfiguration? Validation { get; init; }
}

public sealed class ValidationConfiguration
{
    public bool Enabled { get; init; }
    public string? WindowsCommand { get; init; }
    public bool FullExportEnabled { get; init; }
    public string? FullExportCommand { get; init; }
}

public static class LubanMergeConfigurationLoader
{
    public static MergeCommandOptions Apply(MergeCommandOptions options)
    {
        var configPath = FindConfiguration(options);
        if (configPath is null)
        {
            var defaultRepositoryRoot = Path.GetFullPath(options.RepositoryRoot);
            var defaultValidationEnabled = options.Validate || options.ProjectValidationEnabled;
            var defaultValidationCommand = options.ProjectValidationCommand;
            defaultValidationCommand ??= defaultValidationEnabled
                ? Path.Combine(defaultRepositoryRoot, "ConfigLuban", "check.bat")
                : null;
            if (defaultValidationEnabled && !File.Exists(defaultValidationCommand))
                throw new MergeInputException($"项目快速校验命令不存在：{defaultValidationCommand}。");
            var defaultFullExportEnabled = options.FullExportValidationEnabled;
            var defaultFullExportCommand = options.FullExportValidationCommand;
            defaultFullExportCommand ??= defaultFullExportEnabled
                ? Path.Combine(defaultRepositoryRoot, "ConfigLuban", "gen-pipeline.bat")
                : null;
            if (defaultFullExportEnabled && !File.Exists(defaultFullExportCommand))
                throw new MergeInputException($"完整导出校验命令不存在：{defaultFullExportCommand}。");
            return options with
            {
                RecalculateWithExcel = options.RecalculateWithExcel ?? "never",
                ProjectValidationEnabled = defaultValidationEnabled,
                ProjectValidationCommand = defaultValidationCommand,
                FullExportValidationEnabled = defaultFullExportEnabled,
                FullExportValidationCommand = defaultFullExportCommand
            };
        }

        LubanMergeConfiguration configuration;
        try
        {
            var json = File.ReadAllText(configPath);
            configuration = JsonSerializer.Deserialize<LubanMergeConfiguration>(json, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                PropertyNameCaseInsensitive = false,
                UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
            }) ?? throw new MergeInputException($"配置文件 {configPath} 不能为空。");
        }
        catch (JsonException exception)
        {
            var path = string.IsNullOrWhiteSpace(exception.Path) ? "$" : exception.Path;
            throw new MergeInputException($"配置文件 {configPath} 的 {path} 无效：{exception.Message}", exception);
        }
        catch (IOException exception)
        {
            throw new MergeInputException($"读取配置文件失败：{configPath}。", exception);
        }

        Validate(configuration, configPath);
        var repositoryRoot = Path.GetFullPath(options.RepositoryRoot);
        var validationCommand = ResolveOptionalPath(repositoryRoot, configuration.Validation?.WindowsCommand);
        var validationEnabled = options.Validate || configuration.Validation?.Enabled == true;
        validationCommand ??= validationEnabled
            ? Path.Combine(repositoryRoot, "ConfigLuban", "check.bat")
            : null;
        if (validationEnabled && !File.Exists(validationCommand))
            throw new MergeInputException(
                $"配置文件 {configPath} 的 $.validation.windowsCommand 不存在：{validationCommand}。");
        var fullExportEnabled = options.FullExportValidationEnabled ||
                                configuration.Validation?.FullExportEnabled == true;
        var fullExportCommand = options.FullExportValidationCommand ??
                                ResolveOptionalPath(repositoryRoot, configuration.Validation?.FullExportCommand);
        fullExportCommand ??= fullExportEnabled
            ? Path.Combine(repositoryRoot, "ConfigLuban", "gen-pipeline.bat")
            : null;
        if (fullExportEnabled && !File.Exists(fullExportCommand))
            throw new MergeInputException(
                $"配置文件 {configPath} 的 $.validation.fullExportCommand 不存在：{fullExportCommand}。");
        return options with
        {
            DataRoot = options.DataRoot ?? ResolveOptionalPath(repositoryRoot, configuration.DataRoot),
            TablesPath = options.TablesPath ?? ResolveOptionalPath(repositoryRoot, configuration.TablesFile),
            RecalculateWithExcel = options.RecalculateWithExcel ?? configuration.FormulaRecalculation ?? "never",
            LoadedConfigPath = configPath,
            ProjectValidationEnabled = validationEnabled,
            ProjectValidationCommand = validationCommand,
            ValidateLogicalTableUniqueness = configuration.ValidateLogicalTableUniqueness == true,
            KeyOverrides = configuration.KeyOverrides,
            IgnoredFields = configuration.IgnoredFields,
            InactivePaths = configuration.InactivePaths,
            FullExportValidationEnabled = fullExportEnabled,
            FullExportValidationCommand = fullExportCommand
        };
    }

    private static string? FindConfiguration(MergeCommandOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.ConfigPath))
        {
            var explicitPath = Path.GetFullPath(options.ConfigPath);
            if (!File.Exists(explicitPath))
                throw new MergeInputException($"配置文件不存在：{explicitPath}。");
            return explicitPath;
        }

        var repositoryRoot = Path.GetFullPath(options.RepositoryRoot);
        return new[]
            {
                Path.Combine(repositoryRoot, "luban-excel-merge.json"),
                Path.Combine(repositoryRoot, "ConfigLuban", "luban-excel-merge.json")
            }
            .FirstOrDefault(File.Exists);
    }

    private static void Validate(LubanMergeConfiguration configuration, string path)
    {
        if (configuration.FormulaRecalculation is not null and not ("auto" or "always" or "never"))
            throw new MergeInputException(
                $"配置文件 {path} 的 $.formulaRecalculation 必须是 auto、always 或 never。");
        ValidateInactivePaths(configuration.InactivePaths, path);
        ValidateFieldMappings(configuration.KeyOverrides, path, "keyOverrides", allowEmpty: false);
        ValidateFieldMappings(configuration.IgnoredFields, path, "ignoredFields", allowEmpty: true);
    }

    private static string? ResolveOptionalPath(string repositoryRoot, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        return Path.GetFullPath(value, repositoryRoot);
    }

    private static void ValidateFieldMappings(
        IReadOnlyDictionary<string, string[]>? mappings,
        string path,
        string propertyName,
        bool allowEmpty)
    {
        if (mappings is null)
            return;
        foreach (var mapping in mappings)
        {
            var jsonPath = $"$.{propertyName}['{mapping.Key}']";
            if (string.IsNullOrWhiteSpace(mapping.Key))
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 逻辑表名不能为空。");
            if (mapping.Value is null || (!allowEmpty && mapping.Value.Length == 0))
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 必须包含至少一个字段名。");
            if (mapping.Value.Any(string.IsNullOrWhiteSpace))
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 包含空字段名。");
            if (mapping.Value.Distinct(StringComparer.Ordinal).Count() != mapping.Value.Length)
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 包含重复字段名。");
        }
    }

    private static void ValidateInactivePaths(IReadOnlyList<string>? patterns, string path)
    {
        if (patterns is null)
            return;
        for (var index = 0; index < patterns.Count; index++)
        {
            var pattern = patterns[index];
            var jsonPath = $"$.inactivePaths[{index}]";
            if (string.IsNullOrWhiteSpace(pattern))
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 不能为空。");
            var normalized = pattern.Replace('\\', '/');
            if (Path.IsPathRooted(pattern) || normalized.Split('/').Contains("..", StringComparer.Ordinal))
                throw new MergeInputException($"配置文件 {path} 的 {jsonPath} 必须是仓库内相对路径，且不能包含 ..。");
        }
        if (patterns.Select(PathPatternMatcher.NormalizePattern).Distinct(StringComparer.OrdinalIgnoreCase).Count() != patterns.Count)
            throw new MergeInputException($"配置文件 {path} 的 $.inactivePaths 包含重复模式。");
    }
}
