using System.IO.Compression;
using System.Xml.Linq;
using LubanExcelMerge.Cli;
using LubanExcelMerge.OpenXml;

internal static class CompatibilityRegressionTests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static void Run(string root)
    {
        ColumnMerge(root);
        ExternalTypes(root);
        HeaderGaps(root);
        AuxiliarySheets(root);
    }

    private static void ColumnMerge(string root)
    {
        var fixture = Create(root, "column", "singleton");
        var header = new[] { "##column##var", "##type", "##", "" };
        Write(fixture.BasePath, new[] { header, new[] { "speed", "int", "速度", "10" }, new[] { "hp", "int", "血量", "20" } });
        Write(fixture.LocalPath, new[] { header, new[] { "hp", "int", "本地说明", "20" }, new[] { "speed", "int", "速度", "11" } });
        Write(fixture.RemotePath, new[] { header, new[] { "speed", "int", "速度", "12" }, new[] { "hp", "int", "血量", "25" }, new[] { "new", "string", "新增", "yes" } });
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Assert(session.Conflicts.Count == 1, "column conflict must align by field name after reordering");
        session.Conflicts.Single().Resolve(MergeChoice.Remote);
        session.Save();
        var result = new OpenXmlWorkbookReader().Read(fixture.OutputPath).Sheets.Single();
        Assert(result.GetCell("A1")!.Payload.RawValue == "##column##var", "column marker preserved");
        Assert(result.GetCell("A2")!.Payload.RawValue == "hp" && result.GetCell("D2")!.Payload.RawValue == "25", "remote value on locally reordered row");
        Assert(result.GetCell("C2")!.Payload.RawValue == "本地说明", "local description preserved");
        Assert(result.GetCell("D3")!.Payload.RawValue == "12", "selected value saved");
        Assert(result.Rows.Any(row => row.GetCell(0)?.Payload.RawValue == "new"), "new field preserved");

        Write(fixture.RemotePath, new[] { header, new[] { "speed", "int", "速度", "12" }, new[] { "speed", "int", "重复", "30" } });
        Reject(() => new LubanMergeCoordinator().Prepare(fixture), "duplicate column field names");
        Write(fixture.RemotePath, new[] { new[] { "##column##var", "##", "##type", "" }, new[] { "speed", "速度", "int", "12" } });
        Reject(() => new LubanMergeCoordinator().Prepare(fixture), "changed column layout");
    }

    private static void ExternalTypes(string root)
    {
        var fixture = Create(root, "xml", "map", external: true);
        var defines = Path.Combine(fixture.RepositoryRoot, "Defines");
        Directory.CreateDirectory(defines);
        File.WriteAllText(Path.Combine(defines, "test.xml"), "<module name='legacy'><bean name='Base'><var name='Id' type='int!#range=[0,]'/><bean name='Test'><var name='Value' type='string'/></bean></bean></module>");
        var header = new[] { "##var", "Value", "Id", "#", "#" };
        Write(fixture.BasePath, new[] { header, new[] { "##", "说明", "编号" }, new[] { "", "base", "1", "old-note", "old-note" } });
        Write(fixture.LocalPath, new[] { header, new[] { "##", "说明", "编号" }, new[] { "", "base", "1.0", "old-note", "local-note" } });
        Write(fixture.RemotePath, new[] { header, new[] { "##", "说明", "编号" }, new[] { "", "remote", "1", "remote-note", "old-note" } });
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Assert(session.Conflicts.Count == 0 && session.AddedRecords == 0 && session.DeletedRecords == 0, "XML numeric keys normalize");
        session.Save();
        var sheet = new OpenXmlWorkbookReader().Read(fixture.OutputPath).Sheets.Single();
        Assert(sheet.GetCell("B3")!.Payload.RawValue == "remote", "inherited XML key overrides physical column order");
        Assert(sheet.GetCell("D3")!.Payload.RawValue == "remote-note" && sheet.GetCell("E3")!.Payload.RawValue == "local-note", "duplicate comment columns merge without data loss");
        Write(fixture.RemotePath, new[] { new[] { "##var", "Id", "Unknown" }, new[] { "", "1", "remote" } });
        try { new LubanMergeCoordinator().Prepare(fixture); throw new Exception("unknown XML field accepted"); }
        catch (UnsafeWorkbookException e) { Assert(e.Message.Contains("REMOTE") && e.Message.Contains("Unknown"), "diagnostic identifies side and field"); }
    }

    private static void HeaderGaps(string root)
    {
        var fixture = Create(root, "gaps", "map");
        var rows = new[] { Array.Empty<string>(), new[] { " ##var ", "Id", "Value" }, Array.Empty<string>(), new[] { "##type", "int", "string" }, new[] { "", "1", "base" } };
        Write(fixture.BasePath, rows); Write(fixture.LocalPath, rows);
        Write(fixture.RemotePath, rows.Take(4).Append(new[] { "", "1", "remote" }).ToArray());
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Assert(session.CanSave, "blank header rows accepted");
        session.Save();
        var result = new OpenXmlWorkbookReader().Read(fixture.OutputPath).Sheets.Single();
        Assert(result.GetCell("C5")!.Payload.RawValue == "remote" && result.GetCell("A2")!.Payload.RawValue == " ##var ", "physical row positions and raw markers preserved");
    }

    private static void AuxiliarySheets(string root)
    {
        var fixture = Create(root, "auxiliary", "list");
        foreach (var path in new[] { fixture.BasePath, fixture.LocalPath, fixture.RemotePath })
        {
            TestWorkbookFactory.CreateMultiSheet(path, ("Data", new[] { new[] { "1", "base", "v" } }), ("Notes", Array.Empty<string[]>()));
            ReplaceSheet(path, new[] { new[] { "说明", "不可丢失" } }, "xl/worksheets/sheet2.xml");
        }
        var session = new LubanMergeCoordinator().Prepare(fixture);
        Assert(session.Sheets.Count == 1, "unchanged auxiliary sheet is not parsed as data");
        session.Save();
        Assert(new OpenXmlWorkbookReader().Read(fixture.OutputPath).GetSheet("Notes").GetCell("B1")!.Payload.RawValue == "不可丢失", "auxiliary sheet preserved");
        ReplaceSheet(fixture.RemotePath, new[] { new[] { "说明", "远端修改" } }, "xl/worksheets/sheet2.xml");
        var changed = new LubanMergeCoordinator().Prepare(fixture);
        Assert(!changed.CanSave && changed.Sheets.Single(sheet => sheet.SheetName == "Notes").Conflicts.Count == 1,
            "changed auxiliary sheet must require explicit review rather than be discarded");
    }

    private static MergeCommandOptions Create(string root, string name, string mode, bool external = false)
    {
        var directory = Path.Combine(root, "compatibility", name);
        var data = Path.Combine(directory, "Datas");
        Directory.CreateDirectory(data);
        File.WriteAllText(Path.Combine(data, "__tables__.csv"),
            "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
            $",legacy.TbTest,Test,{(!external).ToString()},Test.xlsx,,{mode},c\n");
        return new MergeCommandOptions(Path.Combine(directory, "base.xlsx"), Path.Combine(directory, "local.xlsx"),
            Path.Combine(directory, "remote.xlsx"), Path.Combine(data, "Test.xlsx"), directory, data, null, null, "zh-CN", false, false, "never", null);
    }

    private static void Write(string path, string[][] rows)
    {
        if (!File.Exists(path)) TestWorkbookFactory.Create(path, Array.Empty<string[]>());
        ReplaceSheet(path, rows);
    }

    private static void ReplaceSheet(string path, string[][] rows, string part = "xl/worksheets/sheet1.xml")
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(part)!.Delete();
        using var stream = archive.CreateEntry(part).Open();
        new XDocument(new XElement(Ns + "worksheet", new XElement(Ns + "sheetData", rows.Select((row, index) =>
            new XElement(Ns + "row", new XAttribute("r", index + 1), row.Select((value, column) =>
                new XElement(Ns + "c", new XAttribute("r", CellReference.Create(index + 1, column)), new XAttribute("t", "inlineStr"),
                    new XElement(Ns + "is", new XElement(Ns + "t", value))))))))).Save(stream);
    }

    private static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
    private static void Reject(Action action, string message)
    {
        try { action(); } catch (UnsafeWorkbookException) { return; }
        throw new Exception("Expected rejection: " + message);
    }
}
