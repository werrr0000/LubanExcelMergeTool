using System.IO.Compression;
using System.Xml.Linq;
using LubanExcelMerge.Core;
using LubanExcelMerge.OpenXml;

internal static class StyleRegressionTests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static IEnumerable<(string Name, Action Run)> Create(string source, string root)
    {
        yield return ("cross-workbook styles remap colliding ids and dependencies", () => Import(source, root, "set", 2));
        yield return ("source default style is remapped across workbooks", () => Import(source, root, "set", 0));
        yield return ("auto recalculation restores valid styles after corruption", () => RecalculationStyles(source, root, WorkbookRecalculationMode.Auto));
        yield return ("always recalculation preserves output on style corruption", () => RecalculationStyles(source, root, WorkbookRecalculationMode.Always));
        yield return ("new remote style is imported and reused", () => Import(source, root, "set", 3));
        yield return ("appended rows import source styles", () => Import(source, root, "append", 3));
        yield return ("inserted rows import source styles", () => Import(source, root, "insert", 3));
        yield return ("metadata rows import source styles", () => Import(source, root, "metadata", 3));
        foreach (var kind in new[] { "cell", "row", "column", "font", "fill", "border", "base", "number" })
        {
            var captured = kind;
            yield return ($"reader rejects invalid {kind} style references", () => RejectInvalid(source, root, captured));
        }
        yield return ("invalid style save preserves output bytes", () => PreserveOutput(source, root));
    }

    private static void Import(string source, string root, string operation, int sourceIndex)
    {
        var prefix = Path.Combine(root, $"style-{operation}-{sourceIndex}");
        var local = prefix + "-local.xlsx";
        var remote = prefix + "-remote.xlsx";
        var output = prefix + "-output.xlsx";
        File.Copy(source, local);
        File.Copy(source, remote);
        Change(local, "xl/styles.xml", doc => doc.Root!.AddFirst(new XElement(Ns + "numFmts",
            new XAttribute("count", 1), new XElement(Ns + "numFmt", new XAttribute("numFmtId", 164), new XAttribute("formatCode", "0.000")))));
        Change(remote, "xl/styles.xml", doc =>
        {
            var styles = doc.Root!;
            styles.AddFirst(new XElement(Ns + "numFmts", new XAttribute("count", 1),
                new XElement(Ns + "numFmt", new XAttribute("numFmtId", 164), new XAttribute("formatCode", "0.00%"))));
            styles.Element(Ns + "fonts")!.Elements().First().Add(new XElement(Ns + "b"));
            styles.Element(Ns + "fills")!.Elements().First().ReplaceWith(new XElement(Ns + "fill",
                new XElement(Ns + "patternFill", new XAttribute("patternType", "solid"), new XElement(Ns + "fgColor", new XAttribute("rgb", "FFFF0000")))));
            styles.Element(Ns + "borders")!.Elements().First().Element(Ns + "left")!.SetAttributeValue("style", "thin");
            var formats = styles.Element(Ns + "cellXfs")!;
            if (sourceIndex == 3)
            {
                formats.Add(new XElement(formats.Elements().ElementAt(2)));
                formats.SetAttributeValue("count", 4);
            }
            var selected = formats.Elements().ElementAt(sourceIndex);
            selected.SetAttributeValue("numFmtId", 164);
            if (selected.Element(Ns + "alignment") is null)
                selected.Add(new XElement(Ns + "alignment", new XAttribute("horizontal", "left")));
        });
        Change(remote, "xl/worksheets/sheet1.xml", doc => doc.Descendants(Ns + "c").Single(c => (string?)c.Attribute("r") == "C4").SetAttributeValue("s", sourceIndex));
        var sourceCell = new OpenXmlWorkbookReader().Read(remote).GetSheet("Data").GetCell("C4")!;
        var cell = new CellWrite(2, sourceCell.Payload, sourceCell.StyleIndex);
        WorkbookEdit[] edits = operation switch
        {
            "append" => [new AppendRowEdit("Data", [cell])],
            "insert" => [new InsertRowEdit("Data", 6, [cell])],
            "metadata" => [new ReplaceMetadataRowsEdit("Data", 1, 1, [new RowWrite([cell])])],
            _ => [new SetCellEdit("Data", "G7", sourceCell.Payload, sourceCell.StyleIndex), new SetCellEdit("Data", "G8", sourceCell.Payload, sourceCell.StyleIndex)]
        };
        var originalLocal = File.ReadAllBytes(local);
        var originalRemote = File.ReadAllBytes(remote);
        new AtomicWorkbookSaver().Save(local, output, edits);
        var address = operation == "metadata" ? "C1" : operation == "set" ? "G7" : "C6";
        var snapshot = new OpenXmlWorkbookReader().Read(output).GetSheet("Data");
        var result = snapshot.GetCell(address)!;
        Assert(result.Payload.ContentEquals(sourceCell.Payload), "cell contents changed");
        if (operation == "set")
            Assert(result.StyleIndex == snapshot.GetCell("G8")!.StyleIndex, "duplicate style import");
        using var archive = ZipFile.OpenRead(output);
        var styles = Read(archive, "xl/styles.xml").Root!;
        var format = styles.Element(Ns + "cellXfs")!.Elements().ElementAt(int.Parse(result.StyleIndex!));
        Assert((string?)format.Attribute("numFmtId") == "165", "custom format collision not remapped");
        Assert(styles.Element(Ns + "numFmts")!.Elements().Single(x => (string?)x.Attribute("numFmtId") == "165").Attribute("formatCode")!.Value == "0.00%", "number format lost");
        Assert(styles.Element(Ns + "fonts")!.Elements().ElementAt((int)format.Attribute("fontId")!).Element(Ns + "b") is not null, "font lost");
        Assert(styles.Element(Ns + "fills")!.Elements().ElementAt((int)format.Attribute("fillId")!).Descendants(Ns + "fgColor").Any(x => (string?)x.Attribute("rgb") == "FFFF0000"), "fill lost");
        Assert((string?)styles.Element(Ns + "borders")!.Elements().ElementAt((int)format.Attribute("borderId")!).Element(Ns + "left")!.Attribute("style") == "thin", "border lost");
        Assert((string?)format.Element(Ns + "alignment")?.Attribute("horizontal") == "left", "alignment lost");
        Assert(File.ReadAllBytes(local).SequenceEqual(originalLocal) && File.ReadAllBytes(remote).SequenceEqual(originalRemote), "input was modified");
    }

    private sealed class CorruptingRecalculator : IWorkbookRecalculator
    {
        public string ProviderName => "Style regression test";
        public bool IsAvailable => true;
        public void Recalculate(string path, TimeSpan timeout) => Change(path, "xl/worksheets/sheet1.xml",
            doc => doc.Descendants(Ns + "c").First().SetAttributeValue("s", 999));
    }

    private static void RecalculationStyles(string source, string root, WorkbookRecalculationMode mode)
    {
        var output = Path.Combine(root, $"recalculated-style-{mode}.xlsx");
        File.Copy(source, output);
        var before = File.ReadAllBytes(output);
        var saver = new AtomicWorkbookSaver(recalculator: new CorruptingRecalculator());
        var edits = new WorkbookEdit[] { new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "updated")) };
        var options = new WorkbookSaveOptions(mode, FormulaMayBeAffected: true);
        if (mode == WorkbookRecalculationMode.Always)
        {
            ExpectInvalid(() => saver.Save(source, output, edits, options));
            Assert(before.SequenceEqual(File.ReadAllBytes(output)), "invalid recalculation replaced output");
        }
        else
        {
            var result = saver.Save(source, output, edits, options);
            Assert(result.RecalculationStatus == WorkbookRecalculationStatus.DeferredAfterRecalculationFailure, "invalid recalculation accepted");
            Assert(new OpenXmlWorkbookReader().Read(output).GetSheet("Data").GetCell("C4")!.Payload.RawValue == "updated", "valid merge was not recovered");
        }
    }

    private static void RejectInvalid(string source, string root, string kind)
    {
        var path = Path.Combine(root, $"invalid-style-{kind}.xlsx");
        File.Copy(source, path);
        if (kind is "cell" or "row" or "column")
            Change(path, "xl/worksheets/sheet1.xml", doc =>
            {
                if (kind == "cell") doc.Descendants(Ns + "c").First().SetAttributeValue("s", 99);
                else if (kind == "row") doc.Descendants(Ns + "row").First().SetAttributeValue("s", 99);
                else doc.Root!.AddFirst(new XElement(Ns + "cols", new XElement(Ns + "col", new XAttribute("min", 1), new XAttribute("max", 1), new XAttribute("style", 99))));
            });
        else
            Change(path, "xl/styles.xml", doc => doc.Root!.Element(Ns + "cellXfs")!.Elements().First().SetAttributeValue(kind switch
            {
                "font" => "fontId", "fill" => "fillId", "border" => "borderId", "base" => "xfId", _ => "numFmtId"
            }, 999));
        ExpectInvalid(() => new OpenXmlWorkbookReader().Read(path));
    }

    private static void PreserveOutput(string source, string root)
    {
        var output = Path.Combine(root, "invalid-style-output.xlsx");
        File.Copy(source, output);
        var original = File.ReadAllBytes(output);
        ExpectInvalid(() => new AtomicWorkbookSaver().Save(source, output,
            [new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "bad"), "999")]));
        Assert(File.ReadAllBytes(output).SequenceEqual(original), "failed save replaced output");
    }

    private static XDocument Read(ZipArchive archive, string part)
    {
        using var stream = archive.GetEntry(part)!.Open();
        return XDocument.Load(stream);
    }

    private static void Change(string path, string part, Action<XDocument> update)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var document = Read(archive, part);
        update(document);
        archive.GetEntry(part)!.Delete();
        using var stream = archive.CreateEntry(part).Open();
        document.Save(stream);
    }

    private static void ExpectInvalid(Action action)
    {
        try { action(); }
        catch (InvalidDataException) { return; }
        throw new InvalidOperationException("Expected invalid style rejection.");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
