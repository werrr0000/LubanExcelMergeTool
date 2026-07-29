using System.Diagnostics;
using System.Text;

namespace LubanExcelMerge.Cli;

public interface IProjectValidator
{
    void Validate(string commandPath, string repositoryRoot, TimeSpan timeout);
}

public sealed class ProjectValidationRunner : IProjectValidator
{
    private readonly string _displayName;

    public ProjectValidationRunner(string displayName = "项目快速校验")
    {
        _displayName = displayName;
    }

    public void Validate(string commandPath, string repositoryRoot, TimeSpan timeout)
    {
        if (!File.Exists(commandPath))
            throw new ProjectValidationException($"{_displayName}命令不存在：{commandPath}。");
        if (timeout <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(timeout));

        var isBatch = Path.GetExtension(commandPath) is ".bat" or ".cmd";
        var startInfo = new ProcessStartInfo
        {
            FileName = isBatch ? Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe" : commandPath,
            Arguments = isBatch ? $"/d /s /c \"\"{commandPath}\"\"" : string.Empty,
            WorkingDirectory = Path.GetDirectoryName(commandPath) ?? repositoryRoot,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        using var process = new Process { StartInfo = startInfo };
        try
        {
            if (!process.Start())
                throw new ProjectValidationException($"无法启动{_displayName}命令。");
            process.StandardInput.WriteLine();
            process.StandardInput.Close();
            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();
            using var cancellation = new CancellationTokenSource(timeout);
            try
            {
                process.WaitForExitAsync(cancellation.Token).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException exception)
            {
                process.Kill(entireProcessTree: true);
                throw new ProjectValidationException(
                    $"{_displayName}在 {timeout.TotalSeconds:0} 秒内未完成。", exception);
            }

            var output = outputTask.GetAwaiter().GetResult();
            var error = errorTask.GetAwaiter().GetResult();
            if (process.ExitCode != 0)
            {
                var details = Tail(string.Join(Environment.NewLine, output, error).Trim(), 4_000);
                throw new ProjectValidationException(
                    $"{_displayName}失败，退出码 {process.ExitCode}。{Environment.NewLine}{details}");
            }
        }
        catch (ProjectValidationException)
        {
            throw;
        }
        catch (Exception exception) when (exception is IOException or InvalidOperationException)
        {
            throw new ProjectValidationException($"运行{_displayName}失败。", exception);
        }
    }

    private static string Tail(string value, int maximumLength) =>
        value.Length <= maximumLength ? value : value[^maximumLength..];
}
