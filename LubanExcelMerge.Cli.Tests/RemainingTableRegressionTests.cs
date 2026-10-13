using System.IO.Compression;
using System.Xml.Linq;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Git;
using LubanExcelMerge.OpenXml;

internal static class RemainingTableRegressionTests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static void PlainHeaders(string root)
    {
        var fixture = Create(root, "plain");
        var prefix = new[] { new[] { "原始文本", "编号", "中文" }, new[] { "SourceText", "Id", "Chinese" } };
        Write(fixture.BasePath, prefix.Concat(new[] { new[] { "Fail", "1", "旧文案" } }).ToArray());
        Write(fixture.LocalPath, prefix.Concat(new[] { new[] { "Fail", "1", "旧文案" }, new[] { "New", "2", "本地新增" } }).ToArray());
        Write(fixture.RemotePath, prefix.Concat(new[] { new[] { "Fail", "1", "远端文案" } }).ToArray());
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Check(session.CanSave && session.KeyName == "Id", "Id must be the key even outside column A");
        session.Save();
        var sheet = new OpenXmlWorkbookReader().Read(fixture.OutputPath).Sheets.Single();
        Check(sheet.GetCell("A2")!.Payload.RawValue == "SourceText" && sheet.GetCell("B2")!.Payload.RawValue == "Id", "header cells remain in original columns");
        Check(sheet.GetCell("C3")!.Payload.RawValue == "远端文案" && sheet.GetCell("B4")!.Payload.RawValue == "2", "remote change and local addition both survive");
    }

    internal static void UnsafePlainHeaders(string root)
    {
        var fixture = Create(root, "unsafe-plain");
        var rows = new[] { new[] { "Id", "Value" }, new[] { "1", "base" }, new[] { "1", "duplicate" } };
        Write(fixture.BasePath, rows); Write(fixture.LocalPath, rows); Write(fixture.RemotePath, rows);
        Reject(() => new LubanMergeCoordinator().Prepare(fixture));
        rows = rows.Take(2).ToArray();
        Write(fixture.BasePath, rows); Write(fixture.LocalPath, rows);
        Write(fixture.RemotePath, new[] { new[] { "Value", "Id" }, new[] { "remote", "1" } });
        Reject(() => new LubanMergeCoordinator().Prepare(fixture));
        Write(fixture.RemotePath, new[] { new[] { "Id", "id" }, new[] { "1", "2" } });
        Reject(() => new LubanMergeCoordinator().Prepare(fixture));
    }

    internal static void AuxiliaryReview(string root)
    {
        var fixture = CreateAuxiliary(root, "aux-review");
        Write(fixture.RemotePath, new[] { new[] { "说明", "远端内容" } }, "xl/worksheets/sheet2.xml");
        var session = new LubanMergeCoordinator().Prepare(fixture);
        var note = session.Sheets.Single(sheet => sheet.SheetName == "Notes");
        Check(note.UsesFixedCellReview && !session.CanSave && note.Conflicts.Count == 1, "all auxiliary content changes must be visible and gate save");
        var conflict = note.Conflicts.Single();
        conflict.Resolve(MergeChoice.Remote);
        Check(session.CanSave && note.Comparison.CreateTable(MergeGridSide.Merged).Rows[0].Cells[1].DisplayValue == "远端内容", "source choice updates preview");
        conflict.ClearResolution();
        Check(!session.CanSave, "undo restores the save gate");
        conflict.Resolve(MergeChoice.Remote); session.Save();
        Check(new OpenXmlWorkbookReader().Read(fixture.OutputPath).GetSheet("Notes").GetCell("B1")!.Payload.RawValue == "远端内容", "selected auxiliary content is written");
    }

    internal static void AuxiliaryLayout(string root)
    {
        var fixture = CreateAuxiliary(root, "aux-layout");
        using (var archive = ZipFile.Open(fixture.RemotePath, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet2.xml")!;
            XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
            xml.Root!.Element(Ns + "sheetData")!.Element(Ns + "row")!.SetAttributeValue("ht", "31");
            entry.Delete(); using var output = archive.CreateEntry("xl/worksheets/sheet2.xml").Open(); xml.Save(output);
        }
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Check(session.CanSave && session.RequiresStructuralChangeConfirmation, "format-only differences need a visible LOCAL layout confirmation");
        Reject(() => new LubanMergeCoordinator().Merge(fixture with { Headless = true }));
        session.Save();
        Check(ReadPart(fixture.LocalPath, "xl/worksheets/sheet2.xml").SequenceEqual(ReadPart(fixture.OutputPath, "xl/worksheets/sheet2.xml")), "auxiliary layout must remain byte-for-byte LOCAL");
    }

    internal static void LfsConflictText(string root)
    {
        var path = Path.Combine(root, "lfs-conflict.xlsx");
        File.WriteAllText(path, "version https://git-lfs.github.com/spec/v1\n<<<<<<< HEAD\noid sha256:local\n=======\noid sha256:remote\n>>>>>>> branch\n");
        try { new GitLfsInputResolver().Resolve(path, root); }
        catch (InvalidDataException e) { Check(e.Message.Contains("冲突文本") && e.Message.Contains("BASE"), "input error explains the cause and recovery"); return; }
        throw new Exception("LFS conflict text was not rejected");
    }

    private static MergeCommandOptions Create(string root, string name)
    {
        var directory = Path.Combine(root, "remaining-tables", name); var data = Path.Combine(directory, "Datas");
        Directory.CreateDirectory(data);
        File.WriteAllText(Path.Combine(data, "__tables__.csv"), "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n,TbTest,Test,TRUE,Test.xlsx,,map,c\n");
        return new MergeCommandOptions(Path.Combine(directory, "base.xlsx"), Path.Combine(directory, "local.xlsx"), Path.Combine(directory, "remote.xlsx"),
            Path.Combine(data, "Test.xlsx"), directory, data, null, null, "zh-CN", false, false, "never", null);
    }

    private static MergeCommandOptions CreateAuxiliary(string root, string name)
    {
        var fixture = Create(root, name);
        foreach (var path in new[] { fixture.BasePath, fixture.LocalPath, fixture.RemotePath })
        {
            TestWorkbookFactory.CreateMultiSheet(path, ("Data", new[] { new[] { "1", "base", "v" } }), ("Notes", Array.Empty<string[]>()));
            Write(path, new[] { new[] { "说明", "原内容" } }, "xl/worksheets/sheet2.xml");
        }
        return fixture;
    }

    private static byte[] ReadPart(string path, string part)
    {
        using var archive = ZipFile.OpenRead(path); using var input = archive.GetEntry(part)!.Open(); using var output = new MemoryStream();
        input.CopyTo(output); return output.ToArray();
    }

    private static void Write(string path, string[][] rows, string part = "xl/worksheets/sheet1.xml")
    {
        if (!File.Exists(path)) TestWorkbookFactory.Create(path, Array.Empty<string[]>());
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update); archive.GetEntry(part)!.Delete();
        using var output = archive.CreateEntry(part).Open();
        new XDocument(new XElement(Ns + "worksheet", new XElement(Ns + "sheetData", rows.Select((row, index) =>
            new XElement(Ns + "row", new XAttribute("r", index + 1), row.Select((value, column) =>
                new XElement(Ns + "c", new XAttribute("r", CellReference.Create(index + 1, column)), new XAttribute("t", "inlineStr"),
                    new XElement(Ns + "is", new XElement(Ns + "t", value))))))))).Save(output);
    }

    private static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    private static void Reject(Action action)
    {
        try { action(); } catch (UnsafeWorkbookException) { return; }
        throw new Exception("Expected safe rejection");
    }
}
