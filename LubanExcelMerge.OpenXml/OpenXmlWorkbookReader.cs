using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using LubanExcelMerge.Core;

namespace LubanExcelMerge.OpenXml;

public sealed class OpenXmlWorkbookReader
{
    private readonly OpenXmlReadLimits _limits;

    public OpenXmlWorkbookReader(OpenXmlReadLimits? limits = null) => _limits = limits ?? new OpenXmlReadLimits();

    public WorkbookSnapshot Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("MVP 只支持 .xlsx 工作簿。");

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ValidateArchive(archive);

        var workbook = LoadXml(archive, "xl/workbook.xml");
        var relationships = LoadXml(archive, "xl/_rels/workbook.xml.rels");
        var relationshipTargets = relationships.Root!
            .Elements(OpenXmlNamespaces.PackageRelationships + "Relationship")
            .Where(element => !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                element => (string)element.Attribute("Id")!,
                element => ResolvePartPath("xl/workbook.xml", (string)element.Attribute("Target")!),
                StringComparer.Ordinal);
        var sharedStrings = ReadSharedStrings(archive);
        var sheets = new List<SheetSnapshot>();

        foreach (var sheetElement in workbook.Root!
                     .Element(OpenXmlNamespaces.Spreadsheet + "sheets")!
                     .Elements(OpenXmlNamespaces.Spreadsheet + "sheet"))
        {
            var name = (string)sheetElement.Attribute("name")!;
            var relationshipId = (string)sheetElement.Attribute(OpenXmlNamespaces.OfficeRelationships + "id")!;
            if (!relationshipTargets.TryGetValue(relationshipId, out var partPath))
                throw new InvalidDataException($"工作表 {name} 的关系 {relationshipId} 无法解析。");

            sheets.Add(ReadSheet(archive, path, name, partPath, sharedStrings));
        }

        return new WorkbookSnapshot(
            Path.GetFullPath(path),
            sheets,
            archive.Entries.Select(entry => entry.FullName).Order(StringComparer.Ordinal).ToArray());
    }

    private void ValidateArchive(ZipArchive archive)
    {
        if (archive.Entries.Count > _limits.MaxEntries)
            throw new InvalidDataException($"工作簿包含 {archive.Entries.Count} 个 ZIP 条目，超过安全上限。");

        long totalLength = 0;
        foreach (var entry in archive.Entries)
        {
            totalLength = checked(totalLength + entry.Length);
            if (totalLength > _limits.MaxTotalUncompressedBytes)
                throw new InvalidDataException("工作簿解压后总大小超过安全上限。");

            if (entry.Length > 0 && entry.CompressedLength == 0)
                throw new InvalidDataException($"ZIP 条目 {entry.FullName} 的压缩长度无效。");
            if (entry.CompressedLength > 0 && entry.Length / (double)entry.CompressedLength > _limits.MaxCompressionRatio)
                throw new InvalidDataException($"ZIP 条目 {entry.FullName} 的压缩比超过安全上限。");
        }
    }

    internal static XDocument LoadXml(ZipArchive archive, string partPath)
    {
        var entry = archive.GetEntry(partPath) ?? throw new InvalidDataException($"工作簿缺少 {partPath}。");
        using var entryStream = entry.Open();
        using var reader = XmlReader.Create(entryStream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
    {
        if (archive.GetEntry("xl/sharedStrings.xml") is null)
            return Array.Empty<string>();

        var document = LoadXml(archive, "xl/sharedStrings.xml");
        return document.Root!
            .Elements(OpenXmlNamespaces.Spreadsheet + "si")
            .Select(item => string.Concat(item.Descendants(OpenXmlNamespaces.Spreadsheet + "t").Select(text => text.Value)))
            .ToArray();
    }

    private static SheetSnapshot ReadSheet(
        ZipArchive archive,
        string workbookPath,
        string sheetName,
        string partPath,
        IReadOnlyList<string> sharedStrings)
    {
        var document = LoadXml(archive, partPath);
        var rows = new List<OpenXmlRowSnapshot>();
        var sheetData = document.Root?.Element(OpenXmlNamespaces.Spreadsheet + "sheetData")
            ?? throw new InvalidDataException($"工作表 {sheetName} 缺少 sheetData。");

        foreach (var rowElement in sheetData.Elements(OpenXmlNamespaces.Spreadsheet + "row"))
        {
            var rowNumber = ParsePositiveInteger((string?)rowElement.Attribute("r"), $"工作表 {sheetName} 的行号");
            var cells = rowElement.Elements(OpenXmlNamespaces.Spreadsheet + "c")
                .Select(cell => ReadCell(cell, workbookPath, sheetName, sharedStrings))
                .ToArray();
            rows.Add(new OpenXmlRowSnapshot(rowNumber, cells));
        }

        return new SheetSnapshot(sheetName, partPath, rows);
    }

    private static OpenXmlCellSnapshot ReadCell(
        XElement cell,
        string workbookPath,
        string sheetName,
        IReadOnlyList<string> sharedStrings)
    {
        var address = (string?)cell.Attribute("r")
            ?? throw new InvalidDataException($"工作表 {sheetName} 包含没有地址的单元格。");
        var (rowNumber, columnIndex) = CellReference.Parse(address);
        var rawDataType = (string?)cell.Attribute("t");
        var formulaElement = cell.Element(OpenXmlNamespaces.Spreadsheet + "f");
        var formula = formulaElement?.Value;
        var value = cell.Element(OpenXmlNamespaces.Spreadsheet + "v")?.Value;
        var kind = CellValueKind.Blank;
        string? rawValue = value;

        if (formula is not null)
        {
            kind = CellValueKind.Formula;
        }
        else
        {
            switch (rawDataType)
            {
                case "s":
                    kind = CellValueKind.String;
                    if (!int.TryParse(value, out var sharedIndex) || sharedIndex < 0 || sharedIndex >= sharedStrings.Count)
                        throw new InvalidDataException($"单元格 {sheetName}!{address} 的共享字符串索引无效。");
                    rawValue = sharedStrings[sharedIndex];
                    break;
                case "inlineStr":
                    kind = CellValueKind.String;
                    rawValue = string.Concat(cell
                        .Element(OpenXmlNamespaces.Spreadsheet + "is")?
                        .Descendants(OpenXmlNamespaces.Spreadsheet + "t")
                        .Select(text => text.Value) ?? Enumerable.Empty<string>());
                    break;
                case "str":
                    kind = CellValueKind.String;
                    break;
                case "b":
                    kind = CellValueKind.Boolean;
                    break;
                case "e":
                    kind = CellValueKind.Error;
                    break;
                default:
                    kind = value is null ? CellValueKind.Blank : CellValueKind.Number;
                    break;
            }
        }

        var payload = new CellPayload(
            kind,
            rawValue,
            formula,
            formula is null ? null : value,
            workbookPath,
            sheetName,
            address,
            rawDataType,
            formula?.Contains('[', StringComparison.Ordinal) == true && formula.Contains(']', StringComparison.Ordinal),
            formulaElement?.Attributes().ToDictionary(
                attribute => attribute.Name.ToString(),
                attribute => attribute.Value,
                StringComparer.Ordinal));
        return new OpenXmlCellSnapshot(address, rowNumber, columnIndex, (string?)cell.Attribute("s"), payload);
    }

    private static int ParsePositiveInteger(string? value, string description)
    {
        if (!int.TryParse(value, out var result) || result < 1)
            throw new InvalidDataException($"{description}无效：{value}。");
        return result;
    }

    internal static string ResolvePartPath(string sourcePart, string target)
    {
        var sourceUri = new Uri("http://package/" + sourcePart, UriKind.Absolute);
        var resolved = new Uri(sourceUri, target);
        return Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
    }
}
