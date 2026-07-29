namespace LubanExcelMerge.Cli;

public static class Program
{
    public static int Main(string[] args)
    {
        var effectiveArguments = args.Length > 0 && args[0] == "merge" && !args.Contains("--log", StringComparer.Ordinal)
            ? args.Concat(new[] { "--log", MergeDiagnosticLogger.CreateDefaultPath() }).ToArray()
            : args;
        return CliApplication.Run(effectiveArguments, Console.Out, Console.Error);
    }
}
