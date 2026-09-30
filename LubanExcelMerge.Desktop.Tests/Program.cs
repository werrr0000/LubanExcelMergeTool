using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Desktop;
using LubanExcelMerge.Git;

if (args.Length > 1 && args[0] == "--fork-driver")
{
    var mode = args[1];
    return AppBuilder.Configure<MergeApp>().UseHeadless(new AvaloniaHeadlessPlatformOptions())
        .StartWithClassicDesktopLifetime(args[2..], lifetime =>
        {
            lifetime.Startup += (_, _) =>
            {
                var clock = Stopwatch.StartNew();
                var started = false;
                DispatcherTimer.Run(() =>
                {
                    if (clock.Elapsed > TimeSpan.FromSeconds(20)) { lifetime.Shutdown(10); return false; }
                    if (lifetime.MainWindow is not MergeWindow window || window.IsBusy || window.Session is null) return true;
                    if (started) return false;
                    started = true;
                    if (mode == "cancel") { Click(window, "CancelMerge"); return false; }
                    var list = window.GetVisualDescendants().OfType<ListBox>().Single();
                    for (var index = 0; index < window.Session.Conflicts.Count; index++)
                    {
                        list.SelectedIndex = index;
                        Click(window, "ChooseRemote");
                    }
                    Find<CheckBox>(window, "Reviewed").IsChecked = true;
                    Click(window, "SaveMerge");
                    return false;
                }, TimeSpan.FromMilliseconds(25));
            };
        });
}

AppBuilder.Configure<MergeApp>().UseHeadless(new AvaloniaHeadlessPlatformOptions()).SetupWithoutStarting();
var root = Path.Combine(Path.GetTempPath(), "Luban Desktop 中文 ' test " + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
var passed = 0;
try
{
    Run("UI cancellation preserves workbook and unmerged index", () =>
    {
        var fixture = CreateFixture("cancel");
        var before = File.ReadAllBytes(fixture.Output);
        var window = Load(fixture.Args);
        Check(window.Session!.RemainingConflicts == 1, "missing conflict");
        Click(window, "CancelMerge");
        Check(window.ResultExitCode == 1, "cancel must fail");
        Check(before.SequenceEqual(File.ReadAllBytes(fixture.Output)), "cancel changed output");
        Check(Git(fixture.Root, "ls-files", "-u").Length > 0, "cancel cleared conflict");
    });
    Run("UI supports filtering, batch resolution, undo and redo", () =>
    {
        var fixture = CreateFixture("navigation");
        var window = Load(fixture.Args);
        var filter = window.GetVisualDescendants().OfType<ComboBox>().Single(c => c.Items.Count == 3);
        Check(window.Session!.RemainingConflicts == 1, "fixture conflict missing");
        Click(window, "ChooseAllRemote");
        Check(window.Session.RemainingConflicts == 0, "batch resolution failed");
        filter.SelectedIndex = 2;
        Check(window.GetVisualDescendants().OfType<ListBox>().Single().ItemCount == 1, "resolved filter failed");
        Click(window, "Undo");
        Check(window.Session.RemainingConflicts == 1, "undo failed");
        Click(window, "Redo");
        Check(window.Session.RemainingConflicts == 0, "redo failed");
    });
    Run("save survives Fork deleting all input temporary files", () =>
    {
        var fixture = CreateFixture("deleted-inputs");
        var window = Load(fixture.Args);
        foreach (var path in Directory.GetFiles(Path.Combine(fixture.Root, "versions")))
            File.Delete(path);
        Click(window, "ChooseRemote");
        Find<CheckBox>(window, "Reviewed").IsChecked = true;
        Pump(window.SaveAsync());
        Check(window.ResultExitCode == 0, "save failed after input cleanup: " + window.ResultExitCode);
        Check(Cell(fixture.Output, "C4") == "remote", "wrong saved conflict");
        Check(Cell(fixture.Output, "D4") == "remote automatic", "automatic edit lost");
        Check(Git(fixture.Root, "ls-files", "-u").Length == 0, "still unmerged");
    });
    foreach (var choice in Enum.GetValues<MergeChoice>())
    {
        Run($"UI {choice} selection saves shared-engine result and stages only MERGED", () =>
        {
            var fixture = CreateFixture(choice.ToString());
            File.WriteAllText(Path.Combine(fixture.Root, "unrelated.txt"), "keep unstaged");
            var window = Load(fixture.Args);
            Check(!Find<Button>(window, "SaveMerge").IsEnabled, "unresolved save allowed");
            Pump(window.SaveAsync());
            Check(Git(fixture.Root, "ls-files", "-u").Length > 0, "unresolved save staged file");
            Click(window, "Choose" + choice);
            Check(window.Session!.CanSave, "choice failed");
            Check(!Find<Button>(window, "SaveMerge").IsEnabled, "review gate missing");
            Find<CheckBox>(window, "Reviewed").IsChecked = true;
            Check(Find<Button>(window, "SaveMerge").IsEnabled, "resolved save disabled");
            Pump(window.SaveAsync());
            Check(window.ResultExitCode == 0, "save failed: " + window.ResultExitCode);
            Check(Git(fixture.Root, "ls-files", "-u").Length == 0, "index still conflicted");
            Check(!Git(fixture.Root, "diff", "--cached", "--name-only").Contains("unrelated"), "unrelated file staged");
            var expected = choice switch { MergeChoice.Base => "base", MergeChoice.Local => "local", _ => "remote" };
            Check(Cell(fixture.Output, "C4") == expected, "wrong conflict value");
            Check(Cell(fixture.Output, "D4") == "remote automatic", "lost automatic merge");
        });
    }
    Run("UI formula warning remains visible after save and blocks further edits", () =>
    {
        var fixture = CreateFixture("formula");
        foreach (var side in new[] { "base", "local", "remote" })
        {
            var path = Path.Combine(fixture.Root, "versions", side + ".xlsx");
            using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument doc;
            using (var input = entry.Open()) doc = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cell = doc.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == "D4");
            cell.Attribute("t")?.Remove();
            cell.ReplaceNodes(new XElement(ns + "f", "1+1"), new XElement(ns + "v", "2"));
            entry.Delete();
            using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            doc.Save(output);
        }
        var window = Load(fixture.Args);
        Click(window, "ChooseRemote");
        Find<CheckBox>(window, "Reviewed").IsChecked = true;
        var save = window.SaveAsync();
        Check(window.IsBusy, "missing save gate");
        window.Close();
        Check(window.IsVisible, "closed during save");
        Pump(save);
        Check(window.ResultExitCode == 0 && window.IsVisible, "formula warning was not retained");
        Check(!Find<CheckBox>(window, "Reviewed").IsEnabled, "completed session can be edited");
        Click(window, "CancelMerge");
        Check(window.ResultExitCode == 0, "completion lost exit status");
    });
    Run("UI multi-sheet save waits for all sheets", () =>
    {
        var fixture = CreateFixture("multi-sheet");
        foreach (var side in new[] { "base", "local", "remote" })
        {
            var path = Path.Combine(fixture.Root, "versions", side + ".xlsx");
            File.Delete(path);
            TestWorkbookFactory.CreateMultiSheet(path,
                ("角色", new[] { new[] { "1", side, "same" } }),
                ("关卡", new[] { new[] { "2", side, "same" } }));
        }
        var window = Load(fixture.Args);
        Click(window, "ChooseRemote");
        Find<CheckBox>(window, "Reviewed").IsChecked = true;
        Check(!Find<Button>(window, "SaveMerge").IsEnabled, "other sheet conflict ignored");
        window.GetVisualDescendants().OfType<ComboBox>().Single(c => c.Items.Count == 2).SelectedIndex = 1;
        Click(window, "ChooseLocal");
        Find<CheckBox>(window, "Reviewed").IsChecked = true;
        Pump(window.SaveAsync());
        Check(window.ResultExitCode == 0, "multi-sheet save failed");
        var workbook = new LubanExcelMerge.OpenXml.OpenXmlWorkbookReader().Read(fixture.Output);
        Check(workbook.GetSheet("角色").GetCell("C4")!.Payload.RawValue == "remote", "sheet 1 lost");
        Check(workbook.GetSheet("关卡").GetCell("C4")!.Payload.RawValue == "local", "sheet 2 lost");
    });
    Run("UI stage failure returns failure and keeps saved workbook", () =>
    {
        var fixture = CreateFixture("stage-failure");
        var window = Load(fixture.Args);
        Click(window, "ChooseRemote");
        Find<CheckBox>(window, "Reviewed").IsChecked = true;
        File.WriteAllText(Path.Combine(fixture.Root, ".git", "index.lock"), "fixture");
        Pump(window.SaveAsync());
        Check(window.ResultExitCode == ExitCodes.WriteValidationFailed, "staging failure reported success");
        Check(Cell(fixture.Output, "C4") == "remote", "saved output lost");
        Check(Git(fixture.Root, "ls-files", "-u").Length > 0, "failed stage cleared conflict");
        window.Close();
    });
    Run("UI invalid input returns documented exit code", () =>
    {
        var window = Load(["--base", "missing.xlsx"]);
        Check(window.ResultExitCode == ExitCodes.InvalidInput, "invalid input exit code");
        Check(!Find<Button>(window, "SaveMerge").IsEnabled, "invalid save enabled");
        window.Close();
    });
    Run("Unix generated Git commands survive spaces, Chinese and single quotes", () =>
    {
        if (OperatingSystem.IsWindows()) return;
        var fixture = CreateFixture("config");
        var executable = Path.Combine(fixture.Root, "合并 ' tool");
        File.WriteAllText(executable, "fixture");
        var configuration = new GitMergeToolConfiguration(executable, fixture.Root);
        configuration.Validate();
        foreach (var command in configuration.BuildLocalConfigurationCommands())
            Process(fixture.Root, "/bin/sh", "-c", command);
        Check(Git(fixture.Root, "config", "mergetool.LubanExcelMerge.cmd").Trim() == configuration.Command, "shell quoting corrupts command");
    });
    Run("Git staging rejects symlinks escaping the project", () =>
    {
        if (OperatingSystem.IsWindows()) return;
        var fixture = CreateFixture("outside");
        var project = Path.Combine(root, "project");
        Directory.CreateDirectory(project);
        var link = Path.Combine(project, "linked");
        Directory.CreateSymbolicLink(link, fixture.Root);
        try
        {
            new GitMergedFileStager().Stage(project, Path.Combine(link, "ConfigLuban", "Datas", "Test.xlsx"));
            throw new Exception("escaped path allowed");
        }
        catch (GitStagingException) { }
    });
    foreach (var mode in new[] { "save", "cancel" })
    {
        Run($"real Git mergetool + Mac launcher + desktop lifetime: {mode}", () =>
        {
            if (OperatingSystem.IsWindows()) return;
            var fixture = CreateFixture("mergetool-" + mode);
            var tools = Path.Combine(fixture.Root, "工具 ' 路径");
            Directory.CreateDirectory(tools);
            var launcher = Path.Combine(tools, "luban-excel-merge");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "luban-excel-merge"), launcher);
            var driver = Path.Combine(tools, "LubanExcelMerge.Desktop");
            File.WriteAllText(driver, "#!/bin/sh\nexec " + Quote(Environment.ProcessPath!) + " --fork-driver " + mode + " \"$@\"\n");
            File.SetUnixFileMode(driver, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            File.SetUnixFileMode(launcher, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            Git(fixture.Root, "config", "mergetool.luban-test.cmd", Quote(launcher) + " \"$BASE\" \"$LOCAL\" \"$REMOTE\" \"$MERGED\"");
            Git(fixture.Root, "config", "mergetool.luban-test.trustExitCode", "true");
            Git(fixture.Root, "config", "mergetool.keepBackup", "false");
            var result = Process(fixture.Root, "git", "mergetool", "--no-prompt", "--tool=luban-test", "--", "ConfigLuban/Datas/Test.xlsx");
            Check((result.Code == 0) == (mode == "save"), "mergetool exit: " + result.Text);
            Check((Git(fixture.Root, "ls-files", "-u").Length == 0) == (mode == "save"), "wrong index after mergetool");
            if (mode == "save") Check(Cell(fixture.Output, "C4") == "remote", "wrong mergetool output");
        });
    }
    Console.WriteLine($"Executed {passed} desktop/integration tests: all passed.");
    return 0;
}
catch (Exception e) { Console.Error.WriteLine(e); return 1; }
finally { Directory.Delete(root, recursive: true); }

void Run(string name, Action test) { test(); passed++; Console.WriteLine("PASS " + name); }
static string Quote(string value) => "'" + value.Replace("'", "'\"'\"'") + "'";
static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
static T Find<T>(MergeWindow window, string name) where T : Control =>
    window.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
static void Click(MergeWindow window, string name) => Find<Button>(window, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
static MergeWindow Load(string[] args)
{
    var window = new MergeWindow(args);
    // 先加载再 Show，避免 Opened 启动第二个异步任务。
    Pump(window.LoadAsync());
    window.Show();
    Dispatcher.UIThread.RunJobs();
    return window;
}
static void Pump(Task task)
{
    var clock = Stopwatch.StartNew();
    while (!task.IsCompleted)
    {
        Dispatcher.UIThread.RunJobs();
        if (clock.Elapsed > TimeSpan.FromSeconds(30)) throw new TimeoutException("UI operation timed out");
        Thread.Sleep(5);
    }
    task.GetAwaiter().GetResult();
    Dispatcher.UIThread.RunJobs();
}
Fixture CreateFixture(string name)
{
    var dir = Path.Combine(root, name);
    Directory.CreateDirectory(dir);
    Git(dir, "init", "-b", "main");
    Git(dir, "config", "user.name", "Luban Test");
    Git(dir, "config", "user.email", "test@example.invalid");
    Git(dir, "config", "commit.gpgsign", "false");
    var data = Path.Combine(dir, "ConfigLuban", "Datas");
    Directory.CreateDirectory(data);
    File.WriteAllText(Path.Combine(data, "__tables__.csv"), "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n##,说明,,,,,,\n,TbTest,Test,TRUE,Test.xlsx,Id,map,c\n");
    var output = Path.Combine(data, "Test.xlsx");
    var versions = Path.Combine(dir, "versions");
    Directory.CreateDirectory(versions);
    string Version(string side, string auto)
    {
        var path = Path.Combine(versions, side + ".xlsx");
        TestWorkbookFactory.Create(path, [new[] { "1", side, auto }]);
        return path;
    }
    var b = Version("base", "unchanged");
    var l = Version("local", "unchanged");
    var r = Version("remote", "remote automatic");
    File.Copy(b, output);
    Git(dir, "add", "ConfigLuban");
    Git(dir, "commit", "-m", "base");
    Git(dir, "checkout", "-b", "remote");
    File.Copy(r, output, true);
    Git(dir, "commit", "-am", "remote");
    Git(dir, "checkout", "main");
    File.Copy(l, output, true);
    Git(dir, "commit", "-am", "local");
    var merge = Process(dir, "git", "merge", "remote");
    Check(merge.Code == 1, "fixture did not conflict");
    return new Fixture(dir, output, ["--base", b, "--local", l, "--remote", r, "--output", output, "--repo-root", dir]);
}
static string Cell(string path, string address)
{
    using var zip = ZipFile.OpenRead(path);
    using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
    var doc = XDocument.Load(stream);
    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    return doc.Descendants(ns + "c").Single(c => (string?)c.Attribute("r") == address).Value;
}
static string Git(string dir, params string[] args)
{
    var result = Process(dir, "git", args);
    if (result.Code != 0) throw new Exception(result.Text);
    return result.Text;
}
static (int Code, string Text) Process(string dir, string file, params string[] args)
{
    var start = new ProcessStartInfo(file) { WorkingDirectory = dir, RedirectStandardOutput = true, RedirectStandardError = true };
    foreach (var arg in args) start.ArgumentList.Add(arg);
    using var process = System.Diagnostics.Process.Start(start)!;
    var output = process.StandardOutput.ReadToEndAsync();
    var error = process.StandardError.ReadToEndAsync();
    if (!process.WaitForExit(30000)) { process.Kill(true); throw new TimeoutException(file); }
    return (process.ExitCode, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
}
record Fixture(string Root, string Output, string[] Args);
