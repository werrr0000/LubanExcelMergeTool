using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using LubanExcelMerge.OpenXml;

internal static class TestWorkbookFactory
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static void Create(string path, IReadOnlyList<string[]> records)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);
        Add(archive, "xl/worksheets/sheet1.xml", CreateSheet(records));
    }

    internal static void CreateWithSchema(
        string path,
        IReadOnlyList<string> fieldNames,
        IReadOnlyList<string> fieldTypes,
        IReadOnlyList<string[]> records)
    {
        if (fieldNames.Count == 0 || fieldNames.Count != fieldTypes.Count)
            throw new ArgumentException("Field names and types must be non-empty and aligned.");
        if (records.Any(record => record.Length != fieldNames.Count))
            throw new ArgumentException("Every record must match the field count.", nameof(records));

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);
        Add(archive, "xl/worksheets/sheet1.xml", CreateSheet(fieldNames, fieldTypes, records));
    }

    internal static void CreateMultiSheet(
        string path,
        params (string Name, IReadOnlyList<string[]> Records)[] sheets)
    {
        if (sheets.Length == 0)
            throw new ArgumentException("At least one sheet is required.", nameof(sheets));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        var contentTypes = new XElement(
            XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Types",
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Default",
                new XAttribute("Extension", "rels"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Default",
                new XAttribute("Extension", "xml"),
                new XAttribute("ContentType", "application/xml")),
            new XElement(
                XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types") + "Override",
                new XAttribute("PartName", "/xl/workbook.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")));
        var relationshipNamespace = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var workbookRelationshipNamespace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        var sheetElements = new List<XElement>();
        var relationshipElements = new List<XElement>();
        for (var index = 0; index < sheets.Length; index++)
        {
            var number = index + 1;
            contentTypes.Add(new XElement(
                contentTypes.Name.Namespace + "Override",
                new XAttribute("PartName", $"/xl/worksheets/sheet{number}.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            sheetElements.Add(new XElement(
                Spreadsheet + "sheet",
                new XAttribute("name", sheets[index].Name),
                new XAttribute("sheetId", number),
                new XAttribute(workbookRelationshipNamespace + "id", $"rId{number}")));
            relationshipElements.Add(new XElement(
                relationshipNamespace + "Relationship",
                new XAttribute("Id", $"rId{number}"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", $"worksheets/sheet{number}.xml")));
            Add(archive, $"xl/worksheets/sheet{number}.xml", CreateSheet(sheets[index].Records));
        }

        Add(archive, "[Content_Types].xml", new XDocument(contentTypes).ToString(SaveOptions.DisableFormatting));
        Add(archive, "_rels/.rels", RootRelationships);
        Add(
            archive,
            "xl/workbook.xml",
            new XDocument(new XElement(
                Spreadsheet + "workbook",
                new XAttribute(XNamespace.Xmlns + "r", workbookRelationshipNamespace),
                new XElement(Spreadsheet + "sheets", sheetElements))).ToString(SaveOptions.DisableFormatting));
        Add(
            archive,
            "xl/_rels/workbook.xml.rels",
            new XDocument(new XElement(relationshipNamespace + "Relationships", relationshipElements))
                .ToString(SaveOptions.DisableFormatting));
    }

    internal static void CreateWithFormula(string path, string value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes);
        Add(archive, "_rels/.rels", RootRelationships);
        Add(archive, "xl/workbook.xml", Workbook);
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships);

        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, "##var", "Id", "A", "B"));
        sheetData.Add(CreateRow(2, "##", "编号", "字段A", "字段B"));
        sheetData.Add(CreateRow(3, "##type", "string", "string", "int"));
        var dataRow = CreateRow(4, string.Empty, "1", value);
        dataRow.Add(new XElement(
            Spreadsheet + "c",
            new XAttribute("r", "D4"),
            new XElement(Spreadsheet + "f", "LEN(C4)"),
            new XElement(Spreadsheet + "v", value.Length)));
        sheetData.Add(dataRow);
        Add(
            archive,
            "xl/worksheets/sheet1.xml",
            new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Spreadsheet + "worksheet", sheetData))
                .ToString(SaveOptions.DisableFormatting));
    }

    private static string CreateSheet(IReadOnlyList<string[]> records)
    {
        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, "##var", "Id", "A", "B"));
        sheetData.Add(CreateRow(2, "##", "编号", "字段A", "字段B"));
        sheetData.Add(CreateRow(3, "##type", "string", "string", "string"));
        for (var index = 0; index < records.Count; index++)
            sheetData.Add(CreateRow(index + 4, string.Empty, records[index][0], records[index][1], records[index][2]));
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Spreadsheet + "worksheet", sheetData)).ToString(SaveOptions.DisableFormatting);
    }

    private static string CreateSheet(
        IReadOnlyList<string> fieldNames,
        IReadOnlyList<string> fieldTypes,
        IReadOnlyList<string[]> records)
    {
        var sheetData = new XElement(Spreadsheet + "sheetData");
        sheetData.Add(CreateRow(1, new[] { "##var" }.Concat(fieldNames).ToArray()));
        sheetData.Add(CreateRow(2, new[] { "##" }.Concat(fieldNames.Select(name =>
            string.IsNullOrEmpty(name) ? string.Empty : $"字段{name}")).ToArray()));
        sheetData.Add(CreateRow(3, new[] { "##type" }.Concat(fieldTypes).ToArray()));
        for (var index = 0; index < records.Count; index++)
        {
            sheetData.Add(CreateRow(
                index + 4,
                new[] { string.Empty }.Concat(records[index]).ToArray()));
        }
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Spreadsheet + "worksheet", sheetData)).ToString(SaveOptions.DisableFormatting);
    }

    private static XElement CreateRow(int rowNumber, params string[] values)
    {
        var row = new XElement(Spreadsheet + "row", new XAttribute("r", rowNumber));
        for (var column = 0; column < values.Length; column++)
        {
            if (values[column].Length == 0)
                continue;
            row.Add(new XElement(
                Spreadsheet + "c",
                new XAttribute("r", CellReference.Create(rowNumber, column)),
                new XAttribute("t", "inlineStr"),
                new XElement(Spreadsheet + "is", new XElement(Spreadsheet + "t", values[column]))));
        }
        return row;
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
          <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
        """;

    private const string WorkbookRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """;
}

