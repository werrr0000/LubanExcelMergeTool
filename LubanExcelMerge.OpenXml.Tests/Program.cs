using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using LubanExcelMerge.Core;
using LubanExcelMerge.Luban;
using LubanExcelMerge.OpenXml;

var testRoot = Path.Combine(Path.GetTempPath(), "LubanExcelMerge.Tests", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testRoot);
try
{
    var sourcePath = Path.Combine(testRoot, "source.xlsx");
    TestWorkbookFactory.Create(sourcePath);
    var tests = new List<(string Name, Action Run)>
    {
        ("reader preserves values, formulas and caches", () => ReaderPreservesPayload(sourcePath)),
        ("reader detects external workbook formulas", () => ReaderDetectsExternalFormula(sourcePath)),
        ("cell references round trip", CellReferencesRoundTrip),
        ("minimal edit preserves style and untouched formula", () => MinimalEditPreservesWorkbook(sourcePath, testRoot)),
        ("row append and delete produce reopenable output", () => AppendAndDeleteRows(sourcePath, testRoot)),
        ("no-op save is byte-for-byte identical", () => NoOpSaveIsByteIdentical(sourcePath, testRoot)),
        ("failed save preserves existing output", () => FailedSavePreservesOutput(sourcePath, testRoot)),
        ("never mode preserves cache and marks full calculation", () => NeverModePreservesCache(sourcePath, testRoot)),
        ("auto mode recalculates only affected formulas", () => AutoModeRecalculatesAffectedFormula(sourcePath, testRoot)),
        ("composite recalculator prefers WPS and falls back to Excel", CompositeRecalculatorSelectsProvider),
        ("always mode recalculates even without edits", () => AlwaysModeRecalculatesWithoutEdits(sourcePath, testRoot)),
        ("missing Excel preserves existing output", () => MissingExcelPreservesOutput(sourcePath, testRoot)),
        ("recalculation failure preserves existing output", () => RecalculationFailurePreservesOutput(sourcePath, testRoot)),
        ("archive safety limits are enforced", () => ArchiveLimitsAreEnforced(sourcePath))
    };

    if (args is ["--real", var accountPath, var battlePath, var formulaPath])
        tests.Add(("real project workbook smoke test", () => RealWorkbookSmokeTest(accountPath, battlePath, formulaPath, testRoot)));

    var failures = new List<string>();
    foreach (var test in tests)
    {
        try
        {
            test.Run();
            Console.WriteLine($"PASS {test.Name}");
        }
        catch (Exception exception)
        {
            failures.Add($"FAIL {test.Name}: {exception.Message}");
        }
    }

    foreach (var failure in failures)
        Console.Error.WriteLine(failure);

    Console.WriteLine($"Executed {tests.Count} tests: {tests.Count - failures.Count} passed, {failures.Count} failed.");
    return failures.Count == 0 ? 0 : 1;
}
finally
{
    var resolvedRoot = Path.GetFullPath(testRoot);
    var resolvedTemp = Path.GetFullPath(Path.GetTempPath());
    if (resolvedRoot.StartsWith(resolvedTemp, StringComparison.OrdinalIgnoreCase) && Directory.Exists(resolvedRoot))
        Directory.Delete(resolvedRoot, recursive: true);
}

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"Expected {expected}, got {actual}.");
}

static void True(bool value)
{
    if (!value)
        throw new InvalidOperationException("Expected true, got false.");
}

static void Throws<TException>(Action action) where TException : Exception
{
    try
    {
        action();
    }
    catch (TException)
    {
        return;
    }

    throw new InvalidOperationException($"Expected {typeof(TException).Name}.");
}

static void ReaderPreservesPayload(string sourcePath)
{
    var snapshot = new OpenXmlWorkbookReader().Read(sourcePath);
    Equal(2, snapshot.Sheets.Count);
    Equal("Data", snapshot.Sheets[0].Name);
    Equal("Name", snapshot.GetSheet("Data").GetCell("C1")!.Payload.RawValue);
    Equal("Local", snapshot.GetSheet("Data").GetCell("C4")!.Payload.RawValue);
    Equal("B5*2", snapshot.GetSheet("Data").GetCell("D5")!.Payload.FormulaText);
    Equal("4", snapshot.GetSheet("Data").GetCell("D5")!.Payload.CachedValue);
    Equal("shared", snapshot.GetSheet("Data").GetCell("E5")!.Payload.FormulaAttributes!["t"]);
    Equal(string.Empty, snapshot.GetSheet("Data").GetCell("E5")!.Payload.FormulaText);
}

static void ReaderDetectsExternalFormula(string sourcePath)
{
    var formula = new OpenXmlWorkbookReader().Read(sourcePath).GetSheet("Data").GetCell("D4")!.Payload;
    True(formula.HasExternalWorkbookReference);
    Equal("10", formula.CachedValue);
}

static void CellReferencesRoundTrip()
{
    foreach (var address in new[] { "A1", "Z9", "AA10", "XFD1048576" })
    {
        var parsed = CellReference.Parse(address);
        Equal(address, CellReference.Create(parsed.RowNumber, parsed.ColumnIndex));
    }
}

static void MinimalEditPreservesWorkbook(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "minimal.xlsx");
    File.Copy(sourcePath, outputPath);
    var sourceHash = HashFile(sourcePath);
    var customPartHash = HashZipPart(sourcePath, "customXml/item1.xml");
    var edits = new WorkbookEdit[]
    {
        new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "Remote Name"))
    };

    var result = new AtomicWorkbookSaver().Save(sourcePath, outputPath, edits);
    var output = new OpenXmlWorkbookReader().Read(outputPath);
    var changed = output.GetSheet("Data").GetCell("C4")!;
    Equal("Remote Name", changed.Payload.RawValue);
    Equal("2", changed.StyleIndex);
    Equal("B5*2", output.GetSheet("Data").GetCell("D5")!.Payload.FormulaText);
    Equal("4", output.GetSheet("Data").GetCell("D5")!.Payload.CachedValue);
    Equal(customPartHash, HashZipPart(outputPath, "customXml/item1.xml"));
    Equal(sourceHash, HashFile(sourcePath));
    Equal(2, result.SheetCount);
    Equal(4, result.FormulaCount);
}

static void AppendAndDeleteRows(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "rows.xlsx");
    var edits = new WorkbookEdit[]
    {
        new SetCellEdit("Data", "D4", new CellPayload(
            CellValueKind.Formula,
            FormulaText: "B4*20",
            CachedValue: "20")),
        new DeleteRowEdit("Data", 5),
        new AppendRowEdit("Data", new[]
        {
            new CellWrite(1, new CellPayload(CellValueKind.Number, "3")),
            new CellWrite(2, new CellPayload(CellValueKind.String, "Added"))
        })
    };

    new AtomicWorkbookSaver().Save(sourcePath, outputPath, edits);
    var output = new OpenXmlWorkbookReader().Read(outputPath).GetSheet("Data");
    Equal("B4*20", output.GetCell("D4")!.Payload.FormulaText);
    Equal("20", output.GetCell("D4")!.Payload.CachedValue);
    Equal("3", output.GetCell("B5")!.Payload.RawValue);
    Equal("Added", output.GetCell("C5")!.Payload.RawValue);
    True(output.GetCell("D5") is null);
}

static void FailedSavePreservesOutput(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "existing.xlsx");
    File.Copy(sourcePath, outputPath);
    var before = HashFile(outputPath);
    Throws<InvalidOperationException>(() => new AtomicWorkbookSaver().Save(
        sourcePath,
        outputPath,
        new WorkbookEdit[] { new SetCellEdit("Missing", "A1", CellPayload.Blank) }));
    Equal(before, HashFile(outputPath));
    Equal(0, Directory.GetFiles(testRoot, ".existing.xlsx.*.tmp.xlsx").Length);
}

static void NoOpSaveIsByteIdentical(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "no-op.xlsx");
    new AtomicWorkbookSaver().Save(sourcePath, outputPath, Array.Empty<WorkbookEdit>());
    Equal(HashFile(sourcePath), HashFile(outputPath));
}

static void NeverModePreservesCache(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "never-formula.xlsx");
    var result = new AtomicWorkbookSaver().Save(
        sourcePath,
        outputPath,
        new WorkbookEdit[]
        {
            new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "Changed input"))
        },
        new WorkbookSaveOptions(WorkbookRecalculationMode.Never, FormulaMayBeAffected: true));

    Equal(WorkbookRecalculationStatus.SourceCachePreservedUnverified, result.RecalculationStatus);
    Equal("4", new OpenXmlWorkbookReader().Read(outputPath).GetSheet("Data").GetCell("D5")!.Payload.CachedValue);
    using var archive = ZipFile.OpenRead(outputPath);
    using var stream = archive.GetEntry("xl/workbook.xml")!.Open();
    var workbook = XDocument.Load(stream);
    XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    var calculation = workbook.Root!.Element(spreadsheet + "calcPr")!;
    Equal("auto", (string?)calculation.Attribute("calcMode"));
    Equal("1", (string?)calculation.Attribute("fullCalcOnLoad"));
    Equal("1", (string?)calculation.Attribute("forceFullCalc"));
}

static void AutoModeRecalculatesAffectedFormula(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "auto-formula.xlsx");
    var recalculator = new FakeWorkbookRecalculator(true, workbookPath =>
        new OpenXmlWorkbookEditor().Apply(
            workbookPath,
            new WorkbookEdit[]
            {
                new SetCellEdit("Data", "D5", new CellPayload(
                    CellValueKind.Formula,
                    FormulaText: "B5*2",
                    CachedValue: "40"))
            }));
    var saver = new AtomicWorkbookSaver(recalculator: recalculator);
    var result = saver.Save(
        sourcePath,
        outputPath,
        new WorkbookEdit[]
        {
            new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "Changed input"))
        },
        new WorkbookSaveOptions(WorkbookRecalculationMode.Auto, FormulaMayBeAffected: true));

    Equal(1, recalculator.CallCount);
    Equal(WorkbookRecalculationStatus.Completed, result.RecalculationStatus);
    Equal("Test Office", result.RecalculationProvider);
    Equal("40", new OpenXmlWorkbookReader().Read(outputPath).GetSheet("Data").GetCell("D5")!.Payload.CachedValue);

    var unaffectedPath = Path.Combine(testRoot, "auto-unaffected.xlsx");
    var unaffected = saver.Save(
        sourcePath,
        unaffectedPath,
        Array.Empty<WorkbookEdit>(),
        new WorkbookSaveOptions(WorkbookRecalculationMode.Auto, FormulaMayBeAffected: false));
    Equal(1, recalculator.CallCount);
    Equal(WorkbookRecalculationStatus.NotNeeded, unaffected.RecalculationStatus);
    Equal(HashFile(sourcePath), HashFile(unaffectedPath));
}

static void CompositeRecalculatorSelectsProvider()
{
    var unavailableWps = new FakeWorkbookRecalculator(false, providerName: "WPS 表格");
    var excel = new FakeWorkbookRecalculator(true, providerName: "Microsoft Excel");
    var fallback = new CompositeWorkbookRecalculator(unavailableWps, excel);
    fallback.Recalculate("unused.xlsx", TimeSpan.FromSeconds(1));
    Equal(0, unavailableWps.CallCount);
    Equal(1, excel.CallCount);
    Equal("Microsoft Excel", fallback.ProviderName);

    var wps = new FakeWorkbookRecalculator(true, providerName: "WPS 表格");
    var preferred = new CompositeWorkbookRecalculator(wps, excel);
    preferred.Recalculate("unused.xlsx", TimeSpan.FromSeconds(1));
    Equal(1, wps.CallCount);
    Equal(1, excel.CallCount);
    Equal("WPS 表格", preferred.ProviderName);
}

static void AlwaysModeRecalculatesWithoutEdits(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "always-formula.xlsx");
    var recalculator = new FakeWorkbookRecalculator(true);
    var result = new AtomicWorkbookSaver(recalculator: recalculator).Save(
        sourcePath,
        outputPath,
        Array.Empty<WorkbookEdit>(),
        new WorkbookSaveOptions(WorkbookRecalculationMode.Always));
    Equal(1, recalculator.CallCount);
    Equal(WorkbookRecalculationStatus.Completed, result.RecalculationStatus);
}

static void MissingExcelPreservesOutput(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "missing-excel.xlsx");
    File.Copy(sourcePath, outputPath);
    var before = HashFile(outputPath);
    var saver = new AtomicWorkbookSaver(recalculator: new FakeWorkbookRecalculator(false));
    Throws<InvalidOperationException>(() => saver.Save(
        sourcePath,
        outputPath,
        new WorkbookEdit[] { new SetCellEdit("Data", "C4", new CellPayload(CellValueKind.String, "Changed")) },
        new WorkbookSaveOptions(WorkbookRecalculationMode.Auto, FormulaMayBeAffected: true)));
    Equal(before, HashFile(outputPath));
    Equal(0, Directory.GetFiles(testRoot, ".missing-excel.xlsx.*.tmp.xlsx").Length);
}

static void RecalculationFailurePreservesOutput(string sourcePath, string testRoot)
{
    var outputPath = Path.Combine(testRoot, "failed-recalculation.xlsx");
    File.Copy(sourcePath, outputPath);
    var before = HashFile(outputPath);
    var recalculator = new FakeWorkbookRecalculator(true, _ => throw new InvalidOperationException("test failure"));
    var saver = new AtomicWorkbookSaver(recalculator: recalculator);
    Throws<InvalidOperationException>(() => saver.Save(
        sourcePath,
        outputPath,
        Array.Empty<WorkbookEdit>(),
        new WorkbookSaveOptions(WorkbookRecalculationMode.Always)));
    Equal(before, HashFile(outputPath));
    Equal(0, Directory.GetFiles(testRoot, ".failed-recalculation.xlsx.*.tmp.xlsx").Length);
}

static void ArchiveLimitsAreEnforced(string sourcePath)
{
    Throws<InvalidDataException>(() => new OpenXmlWorkbookReader(new OpenXmlReadLimits(MaxEntries: 1)).Read(sourcePath));
}

static void RealWorkbookSmokeTest(
    string accountPath,
    string battlePath,
    string formulaPath,
    string testRoot)
{
    var reader = new OpenXmlWorkbookReader();
    var account = reader.Read(accountPath);
    var accountSchema = ParseSchema(account.Sheets[0]);
    Equal(4, accountSchema.DataStartRowNumber);
    Equal("Id", accountSchema.Fields[0].Name);

    var battle = reader.Read(battlePath);
    var battleSchema = ParseSchema(battle.Sheets[0]);
    Equal(5, battleSchema.DataStartRowNumber);
    True(battleSchema.FindField("Id") is not null);
    True(battleSchema.FindField("SmallGameType") is not null);

    var formula = reader.Read(formulaPath);
    True(formula.Sheets.Sum(sheet => sheet.FormulaCount) > 0);
    var noOpOutput = Path.Combine(testRoot, "real-no-op.xlsx");
    new AtomicWorkbookSaver().Save(formulaPath, noOpOutput, Array.Empty<WorkbookEdit>());
    Equal(HashFile(formulaPath), HashFile(noOpOutput));

    var formulaSheet = formula.Sheets.First(sheet => sheet.FormulaCount > 0);
    var unchangedCell = formulaSheet.Rows.SelectMany(row => row.Cells)
        .First(cell => cell.Payload.Kind is not CellValueKind.Formula and not CellValueKind.Blank);
    var touchedOutput = Path.Combine(testRoot, "real-formula-touched.xlsx");
    var formulaSourceHash = HashFile(formulaPath);
    new AtomicWorkbookSaver().Save(
        formulaPath,
        touchedOutput,
        new WorkbookEdit[] { new SetCellEdit(formulaSheet.Name, unchangedCell.Address, unchangedCell.Payload) });
    var touched = reader.Read(touchedOutput);
    Equal(formula.Sheets.Sum(sheet => sheet.FormulaCount), touched.Sheets.Sum(sheet => sheet.FormulaCount));
    Equal(formulaSourceHash, HashFile(formulaPath));
}

static LubanSchema ParseSchema(SheetSnapshot sheet)
{
    var rows = sheet.Rows.Select(row =>
    {
        var maxColumn = row.Cells.Select(cell => cell.ColumnIndex).DefaultIfEmpty(-1).Max();
        var values = new string?[maxColumn + 1];
        foreach (var cell in row.Cells)
            values[cell.ColumnIndex] = cell.Payload.RawValue;
        return new LubanRawRow(row.RowNumber, values);
    }).ToArray();
    return LubanSchemaParser.Parse(rows);
}

static string HashFile(string path)
{
    using var stream = File.OpenRead(path);
    return Convert.ToHexString(SHA256.HashData(stream));
}

static string HashZipPart(string path, string partName)
{
    using var archive = ZipFile.OpenRead(path);
    using var stream = archive.GetEntry(partName)!.Open();
    return Convert.ToHexString(SHA256.HashData(stream));
}

internal static class TestWorkbookFactory
{
    internal static void Create(string path)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);
        Add(archive, "xl/sharedStrings.xml", SharedStrings);
        Add(archive, "xl/styles.xml", Styles);
        Add(archive, "xl/worksheets/sheet1.xml", DataSheet);
        Add(archive, "xl/worksheets/sheet2.xml", NotesSheet);
        Add(archive, "customXml/item1.xml", "<fixture preserve=\"true\">unchanged</fixture>");
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private const string ContentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        </Types>
        """;

    private const string RootRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private const string Workbook = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Data" sheetId="1" r:id="rId1"/>
            <sheet name="Notes" sheetId="2" r:id="rId2"/>
          </sheets>
          <calcPr calcId="191029" fullCalcOnLoad="0"/>
        </workbook>
        """;

    private const string WorkbookRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
          <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private const string SharedStrings = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8">
          <si><t>##var</t></si><si><t>Id</t></si><si><t>Name</t></si><si><t>Formula</t></si>
          <si><t>Local</t></si><si><t>Keep</t></si><si><t>##</t></si><si><t>##type</t></si>
        </sst>
        """;

    private const string Styles = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf></cellXfs>
        </styleSheet>
        """;

    private const string DataSheet = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>
            <row r="2"><c r="A2" t="s"><v>6</v></c></row>
            <row r="3"><c r="A3" t="s"><v>7</v></c><c r="B3" t="str"><v>int</v></c><c r="C3" t="str"><v>string</v></c><c r="D3" t="str"><v>int</v></c></row>
            <row r="4"><c r="B4"><v>1</v></c><c r="C4" s="2" t="s"><v>4</v></c><c r="D4"><f>XLOOKUP(B4,[Other.xlsx]Sheet1!A:A,[Other.xlsx]Sheet1!B:B)</f><v>10</v></c><c r="E4"><f t="shared" si="0" ref="E4:E5">B4*3</f><v>3</v></c></row>
            <row r="5"><c r="B5"><v>2</v></c><c r="C5" t="s"><v>5</v></c><c r="D5"><f>B5*2</f><v>4</v></c><c r="E5"><f t="shared" si="0"/><v>6</v></c></row>
          </sheetData>
        </worksheet>
        """;

    private const string NotesSheet = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>preserve sheet</t></is></c></row></sheetData>
        </worksheet>
        """;
}

internal sealed class FakeWorkbookRecalculator : IWorkbookRecalculator
{
    private readonly Action<string>? _recalculate;

    internal FakeWorkbookRecalculator(
        bool isAvailable,
        Action<string>? recalculate = null,
        string providerName = "Test Office")
    {
        IsAvailable = isAvailable;
        _recalculate = recalculate;
        ProviderName = providerName;
    }

    public string ProviderName { get; }
    public bool IsAvailable { get; }
    public int CallCount { get; private set; }

    public void Recalculate(string workbookPath, TimeSpan timeout)
    {
        CallCount++;
        _recalculate?.Invoke(workbookPath);
    }
}
