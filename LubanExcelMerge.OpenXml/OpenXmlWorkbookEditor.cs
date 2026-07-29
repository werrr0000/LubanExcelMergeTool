using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using LubanExcelMerge.Core;

namespace LubanExcelMerge.OpenXml;

public sealed record WorkbookEditResult(IReadOnlySet<string> TouchedPartPaths);

public sealed class OpenXmlWorkbookEditor
{
    public WorkbookEditResult Apply(string workbookPath, IReadOnlyList<WorkbookEdit> edits)
    {
        using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        var sheetParts = ReadSheetParts(archive);
        var touchedParts = new HashSet<string>(StringComparer.Ordinal);

        foreach (var editGroup in edits.GroupBy(edit => edit.SheetName, StringComparer.Ordinal))
        {
            if (!sheetParts.TryGetValue(editGroup.Key, out var partPath))
                throw new InvalidOperationException($"工作簿中不存在工作表 {editGroup.Key}。");

            var document = OpenXmlWorkbookReader.LoadXml(archive, partPath);
            var sheetData = document.Root?.Element(OpenXmlNamespaces.Spreadsheet + "sheetData")
                ?? throw new InvalidDataException($"工作表 {editGroup.Key} 缺少 sheetData。");

            foreach (var edit in editGroup)
            {
                switch (edit)
                {
                    case SetCellEdit setCell:
                        SetCell(sheetData, setCell.Address, setCell.Payload, setCell.StyleIndex);
                        break;
                    case DeleteRowEdit deleteRow:
                        DeleteRow(sheetData, deleteRow.RowNumber);
                        break;
                    case AppendRowEdit appendRow:
                        AppendRow(sheetData, appendRow.Cells);
                        break;
                    default:
                        throw new NotSupportedException($"不支持的工作簿编辑类型：{edit.GetType().Name}。");
                }
            }

            ExpandDimension(document, sheetData);
            ReplaceXmlEntry(archive, partPath, document);
            touchedParts.Add(partPath);
        }

        return new WorkbookEditResult(touchedParts);
    }

    public void MarkForFullCalculation(string workbookPath)
    {
        using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: false);
        const string workbookPart = "xl/workbook.xml";
        var document = OpenXmlWorkbookReader.LoadXml(archive, workbookPart);
        var root = document.Root ?? throw new InvalidDataException("工作簿缺少 workbook 根元素。");
        var calculation = root.Element(OpenXmlNamespaces.Spreadsheet + "calcPr");
        if (calculation is null)
        {
            calculation = new XElement(OpenXmlNamespaces.Spreadsheet + "calcPr");
            root.Add(calculation);
        }

        calculation.SetAttributeValue("calcMode", "auto");
        calculation.SetAttributeValue("fullCalcOnLoad", "1");
        calculation.SetAttributeValue("forceFullCalc", "1");
        ReplaceXmlEntry(archive, workbookPart, document);
    }

    private static IReadOnlyDictionary<string, string> ReadSheetParts(ZipArchive archive)
    {
        var workbook = OpenXmlWorkbookReader.LoadXml(archive, "xl/workbook.xml");
        var relationships = OpenXmlWorkbookReader.LoadXml(archive, "xl/_rels/workbook.xml.rels");
        var targets = relationships.Root!
            .Elements(OpenXmlNamespaces.PackageRelationships + "Relationship")
            .Where(element => !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                element => (string)element.Attribute("Id")!,
                element => OpenXmlWorkbookReader.ResolvePartPath("xl/workbook.xml", (string)element.Attribute("Target")!),
                StringComparer.Ordinal);

        return workbook.Root!
            .Element(OpenXmlNamespaces.Spreadsheet + "sheets")!
            .Elements(OpenXmlNamespaces.Spreadsheet + "sheet")
            .ToDictionary(
                sheet => (string)sheet.Attribute("name")!,
                sheet => targets[(string)sheet.Attribute(OpenXmlNamespaces.OfficeRelationships + "id")!],
                StringComparer.Ordinal);
    }

    private static void SetCell(XElement sheetData, string address, CellPayload payload, string? styleIndex)
    {
        var (rowNumber, columnIndex) = CellReference.Parse(address);
        var row = GetOrCreateRow(sheetData, rowNumber);
        var cell = row.Elements(OpenXmlNamespaces.Spreadsheet + "c")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("r"), address, StringComparison.OrdinalIgnoreCase));
        if (cell is null)
        {
            cell = new XElement(OpenXmlNamespaces.Spreadsheet + "c", new XAttribute("r", address));
            InsertCellInOrder(row, cell, columnIndex);
        }

        if (styleIndex is not null)
            cell.SetAttributeValue("s", styleIndex);
        WritePayload(cell, payload);
    }

    private static void DeleteRow(XElement sheetData, int rowNumber)
    {
        if (rowNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));
        sheetData.Elements(OpenXmlNamespaces.Spreadsheet + "row")
            .FirstOrDefault(row => (string?)row.Attribute("r") == rowNumber.ToString())?
            .Remove();
    }

    private static void AppendRow(XElement sheetData, IReadOnlyList<CellWrite> cells)
    {
        var rowNumber = sheetData.Elements(OpenXmlNamespaces.Spreadsheet + "row")
            .Select(row => int.TryParse((string?)row.Attribute("r"), out var number) ? number : 0)
            .DefaultIfEmpty(0)
            .Max() + 1;
        var row = GetOrCreateRow(sheetData, rowNumber);
        foreach (var cell in cells.OrderBy(cell => cell.ColumnIndex))
            SetCell(sheetData, CellReference.Create(rowNumber, cell.ColumnIndex), cell.Payload, cell.StyleIndex);
    }

    private static XElement GetOrCreateRow(XElement sheetData, int rowNumber)
    {
        var existing = sheetData.Elements(OpenXmlNamespaces.Spreadsheet + "row")
            .FirstOrDefault(row => (string?)row.Attribute("r") == rowNumber.ToString());
        if (existing is not null)
            return existing;

        var row = new XElement(OpenXmlNamespaces.Spreadsheet + "row", new XAttribute("r", rowNumber));
        var next = sheetData.Elements(OpenXmlNamespaces.Spreadsheet + "row")
            .FirstOrDefault(candidate => int.TryParse((string?)candidate.Attribute("r"), out var number) && number > rowNumber);
        if (next is null)
            sheetData.Add(row);
        else
            next.AddBeforeSelf(row);
        return row;
    }

    private static void InsertCellInOrder(XElement row, XElement cell, int columnIndex)
    {
        var next = row.Elements(OpenXmlNamespaces.Spreadsheet + "c")
            .FirstOrDefault(candidate =>
            {
                var reference = (string?)candidate.Attribute("r");
                return reference is not null && CellReference.Parse(reference).ColumnIndex > columnIndex;
            });
        if (next is null)
            row.Add(cell);
        else
            next.AddBeforeSelf(cell);
    }

    private static void WritePayload(XElement cell, CellPayload payload)
    {
        cell.Elements()
            .Where(element => element.Name == OpenXmlNamespaces.Spreadsheet + "f" ||
                              element.Name == OpenXmlNamespaces.Spreadsheet + "v" ||
                              element.Name == OpenXmlNamespaces.Spreadsheet + "is")
            .Remove();

        var content = new List<XElement>();
        switch (payload.Kind)
        {
            case CellValueKind.Blank:
                cell.SetAttributeValue("t", null);
                break;
            case CellValueKind.String:
                cell.SetAttributeValue("t", "inlineStr");
                var text = new XElement(OpenXmlNamespaces.Spreadsheet + "t", payload.RawValue ?? string.Empty);
                if (NeedsPreservedWhitespace(payload.RawValue))
                    text.SetAttributeValue(OpenXmlNamespaces.Xml + "space", "preserve");
                content.Add(new XElement(OpenXmlNamespaces.Spreadsheet + "is", text));
                break;
            case CellValueKind.Number:
                cell.SetAttributeValue("t", null);
                content.Add(new XElement(OpenXmlNamespaces.Spreadsheet + "v", payload.RawValue ?? string.Empty));
                break;
            case CellValueKind.Boolean:
                cell.SetAttributeValue("t", "b");
                content.Add(new XElement(OpenXmlNamespaces.Spreadsheet + "v", payload.RawValue ?? "0"));
                break;
            case CellValueKind.Error:
                cell.SetAttributeValue("t", "e");
                content.Add(new XElement(OpenXmlNamespaces.Spreadsheet + "v", payload.RawValue ?? "#VALUE!"));
                break;
            case CellValueKind.Formula:
                if (payload.FormulaText is null)
                    throw new InvalidOperationException("公式单元格缺少公式文本。");
                cell.SetAttributeValue("t", payload.RawDataType);
                var formula = new XElement(OpenXmlNamespaces.Spreadsheet + "f", payload.FormulaText);
                if (payload.FormulaAttributes is not null)
                {
                    foreach (var attribute in payload.FormulaAttributes)
                        formula.SetAttributeValue(XName.Get(attribute.Key), attribute.Value);
                }
                content.Add(formula);
                if (payload.CachedValue is not null)
                    content.Add(new XElement(OpenXmlNamespaces.Spreadsheet + "v", payload.CachedValue));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(payload));
        }

        if (content.Count > 0)
            cell.AddFirst(content);
    }

    private static bool NeedsPreservedWhitespace(string? value) =>
        !string.IsNullOrEmpty(value) && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));

    private static void ExpandDimension(XDocument document, XElement sheetData)
    {
        var cells = sheetData.Descendants(OpenXmlNamespaces.Spreadsheet + "c")
            .Select(cell => (string?)cell.Attribute("r"))
            .Where(address => address is not null)
            .Select(address => CellReference.Parse(address!))
            .ToArray();
        if (cells.Length == 0)
            return;

        var minimumRow = cells.Min(cell => cell.RowNumber);
        var minimumColumn = cells.Min(cell => cell.ColumnIndex);
        var maximumRow = cells.Max(cell => cell.RowNumber);
        var maximumColumn = cells.Max(cell => cell.ColumnIndex);
        var dimension = document.Root?.Element(OpenXmlNamespaces.Spreadsheet + "dimension");
        if (dimension is not null)
        {
            var references = ((string?)dimension.Attribute("ref") ?? string.Empty).Split(':');
            try
            {
                var first = CellReference.Parse(references[0]);
                var last = CellReference.Parse(references.Length > 1 ? references[1] : references[0]);
                minimumRow = Math.Min(minimumRow, first.RowNumber);
                minimumColumn = Math.Min(minimumColumn, first.ColumnIndex);
                maximumRow = Math.Max(maximumRow, last.RowNumber);
                maximumColumn = Math.Max(maximumColumn, last.ColumnIndex);
            }
            catch (FormatException)
            {
                // Replace an invalid dimension with the range proven by actual cells.
            }
        }
        else
        {
            dimension = new XElement(OpenXmlNamespaces.Spreadsheet + "dimension");
            document.Root?.AddFirst(dimension);
        }

        var firstAddress = CellReference.Create(minimumRow, minimumColumn);
        var lastAddress = CellReference.Create(maximumRow, maximumColumn);
        dimension.SetAttributeValue("ref", firstAddress == lastAddress ? firstAddress : $"{firstAddress}:{lastAddress}");
    }

    private static void ReplaceXmlEntry(ZipArchive archive, string partPath, XDocument document)
    {
        archive.GetEntry(partPath)?.Delete();
        var entry = archive.CreateEntry(partPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        document.Save(writer, SaveOptions.DisableFormatting);
    }
}
