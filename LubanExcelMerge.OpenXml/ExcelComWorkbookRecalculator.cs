using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace LubanExcelMerge.OpenXml;

public sealed class ExcelComWorkbookRecalculator : IWorkbookRecalculator
{
    private const int AutomationSecurityForceDisable = 3;
    private const int CalculationDone = 0;
    private readonly IReadOnlyList<string> _programmaticIds;
    private readonly IReadOnlyList<string> _processNames;

    public ExcelComWorkbookRecalculator() : this(
        "Microsoft Excel",
        new[] { "Excel.Application" },
        new[] { "EXCEL" })
    {
    }

    internal ExcelComWorkbookRecalculator(
        string providerName,
        IReadOnlyList<string> programmaticIds,
        IReadOnlyList<string> processNames)
    {
        ProviderName = providerName;
        _programmaticIds = programmaticIds;
        _processNames = processNames;
    }

    public string ProviderName { get; }
    public bool IsAvailable => OperatingSystem.IsWindows() && ResolveApplicationType() is not null;

    public void Recalculate(string workbookPath, TimeSpan timeout)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("待重算的工作簿不存在。", workbookPath);
        if (timeout <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(timeout));
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException($"{ProviderName} COM 重算仅支持 Windows。");

        var applicationType = ResolveApplicationType()
            ?? throw new InvalidOperationException($"本机未安装或未正确注册 {ProviderName} 自动化接口。");
        var existingProcessIds = GetProcessIds(_processNames);
        using var completed = new ManualResetEventSlim();
        Exception? failure = null;
        var applicationProcessId = 0;
        var ownsApplicationProcess = false;

        var worker = new Thread(() =>
        {
            object? application = null;
            object? workbooks = null;
            object? workbook = null;
            try
            {
                application = Activator.CreateInstance(applicationType)
                    ?? throw new InvalidOperationException($"无法启动 {ProviderName}。");
                dynamic office = application;
                var processId = ResolveApplicationProcessId(office, existingProcessIds);
                if (processId <= 0 || existingProcessIds.Contains(processId))
                    throw new InvalidOperationException(
                        $"{ProviderName} 自动化接口复用了已有进程；为保护用户已打开的文档，本次重算已停止。");
                Volatile.Write(ref applicationProcessId, processId);
                ownsApplicationProcess = true;

                SetProperty(application, "Visible", false, required: true);
                SetProperty(application, "DisplayAlerts", false, required: true);
                SetProperty(application, "AskToUpdateLinks", false, required: false);
                SetProperty(application, "EnableEvents", false, required: false);
                SetProperty(application, "AutomationSecurity", AutomationSecurityForceDisable, required: false);

                workbooks = office.Workbooks;
                dynamic books = workbooks;
                workbook = books.Open(
                    Path.GetFullPath(workbookPath),
                    UpdateLinks: 0,
                    ReadOnly: false,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                StartFullCalculation(office);
                var deadline = DateTime.UtcNow + timeout;
                while (!IsCalculationComplete(office))
                {
                    if (DateTime.UtcNow >= deadline)
                        throw new TimeoutException($"{ProviderName} 在 {timeout.TotalSeconds:0} 秒内未完成重算。");
                    Thread.Sleep(100);
                }

                ((dynamic)workbook).Save();
            }
            catch (Exception exception)
            {
                failure = exception;
            }
            finally
            {
                TryClose(workbook);
                ReleaseComObject(workbook);
                ReleaseComObject(workbooks);
                if (ownsApplicationProcess)
                    TryQuit(application);
                ReleaseComObject(application);
                completed.Set();
            }
        })
        {
            IsBackground = true,
            Name = $"LubanExcelMerge.{ProviderName}.Recalculation"
        };
        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();

        if (!completed.Wait(timeout + TimeSpan.FromSeconds(5)))
        {
            KillOwnedProcess(Volatile.Read(ref applicationProcessId));
            throw new TimeoutException(
                $"{ProviderName} 重算超时（{timeout.TotalSeconds:0} 秒），已终止本工具启动的进程。");
        }

        if (failure is not null)
            throw new InvalidOperationException($"{ProviderName} 完整重算失败。", failure);
    }

    private Type? ResolveApplicationType()
    {
        if (!OperatingSystem.IsWindows())
            return null;
        foreach (var programmaticId in _programmaticIds)
        {
            var type = Type.GetTypeFromProgID(programmaticId, throwOnError: false);
            if (type is not null)
                return type;
        }
        return null;
    }

    private int ResolveApplicationProcessId(dynamic application, IReadOnlySet<int> existingProcessIds)
    {
        try
        {
            var windowHandle = Convert.ToInt32(application.Hwnd);
            if (windowHandle != 0)
            {
                _ = GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
                return checked((int)processId);
            }
        }
        catch (Exception exception) when (exception is COMException or RuntimeBinderException)
        {
            // Some WPS editions do not expose Application.Hwnd.
        }

        var newProcessIds = GetProcessIds(_processNames).Except(existingProcessIds).ToArray();
        return newProcessIds.Length == 1 ? newProcessIds[0] : 0;
    }

    private static HashSet<int> GetProcessIds(IEnumerable<string> processNames)
    {
        var result = new HashSet<int>();
        foreach (var processName in processNames)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                using (process)
                    result.Add(process.Id);
            }
        }
        return result;
    }

    private static void StartFullCalculation(dynamic application)
    {
        try
        {
            application.CalculateFullRebuild();
        }
        catch (Exception exception) when (exception is COMException or RuntimeBinderException)
        {
            application.Calculate();
        }
    }

    private static bool IsCalculationComplete(dynamic application)
    {
        try
        {
            return Convert.ToInt32(application.CalculationState) == CalculationDone;
        }
        catch (Exception exception) when (exception is COMException or RuntimeBinderException)
        {
            // WPS Calculate is synchronous in editions without CalculationState.
            return true;
        }
    }

    private static void SetProperty(object target, string name, object value, bool required)
    {
        try
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty,
                null,
                target,
                new[] { value });
        }
        catch (Exception) when (!required)
        {
            // Optional compatibility property is unavailable in this office suite.
        }
    }

    private static void KillOwnedProcess(int processId)
    {
        if (processId <= 0)
            return;
        try
        {
            using var process = Process.GetProcessById(processId);
            process.Kill(entireProcessTree: true);
            process.WaitForExit(5_000);
        }
        catch (ArgumentException)
        {
            // The office process already exited.
        }
    }

    private static void TryClose(object? workbook)
    {
        if (workbook is null)
            return;
        try { ((dynamic)workbook).Close(SaveChanges: false); }
        catch (Exception) { }
    }

    private static void TryQuit(object? application)
    {
        if (application is null)
            return;
        try { ((dynamic)application).Quit(); }
        catch (Exception) { }
    }

    private static void ReleaseComObject(object? value)
    {
        if (!OperatingSystem.IsWindows())
            return;
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);
}
