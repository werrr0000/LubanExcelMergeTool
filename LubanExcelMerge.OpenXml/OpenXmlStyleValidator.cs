using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace LubanExcelMerge.OpenXml;

internal sealed class OpenXmlStyleValidator
{
    private static readonly XNamespace Ns = OpenXmlNamespaces.Spreadsheet;
    private readonly int _cellStyleCount;
    private readonly string _workbookPath;

    internal OpenXmlStyleValidator(ZipArchive archive, string workbookPath)
    {
        _workbookPath = workbookPath;
        if (archive.GetEntry("xl/styles.xml") is null)
        {
            _cellStyleCount = 1;
            return;
        }
        var root = OpenXmlWorkbookReader.LoadXml(archive, "xl/styles.xml").Root
            ?? throw new InvalidDataException($"{workbookPath}：样式表缺少根元素。");
        var cellFormats = root.Element(Ns + "cellXfs")?.Elements(Ns + "xf").ToArray()
            ?? throw new InvalidDataException($"{workbookPath}：样式表缺少 cellXfs。");
        _cellStyleCount = cellFormats.Length;
        var baseFormats = root.Element(Ns + "cellStyleXfs")?.Elements(Ns + "xf").ToArray() ?? [];
        var customFormats = new HashSet<int>();
        foreach (var format in root.Element(Ns + "numFmts")?.Elements(Ns + "numFmt") ?? [])
        {
            var id = Parse((string?)format.Attribute("numFmtId"), "numFmtId");
            if (!customFormats.Add(id) || format.Attribute("formatCode") is null)
                throw new InvalidDataException($"{workbookPath}：自定义数字格式 {id} 重复或缺少格式定义。");
        }
        foreach (var format in cellFormats.Concat(baseFormats))
        {
            foreach (var (attribute, collection) in new[] { ("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders") })
            {
                if (format.Attribute(attribute) is { } reference)
                    ValidateReference(reference.Value, root.Element(Ns + collection)?.Elements().Count() ?? 0, $"样式 {attribute}");
            }
            if (format.Attribute("numFmtId") is { } numberFormat)
            {
                var id = Parse(numberFormat.Value, "numFmtId");
                if (id >= 164 && !customFormats.Contains(id))
                    throw new InvalidDataException($"{workbookPath}：样式引用了缺失的自定义数字格式 {id}。");
            }
        }
        foreach (var format in cellFormats)
            if (format.Attribute("xfId") is { } reference)
                ValidateReference(reference.Value, baseFormats.Length, "样式 xfId");
        foreach (var style in root.Element(Ns + "cellStyles")?.Elements(Ns + "cellStyle") ?? [])
            if (style.Attribute("xfId") is { } reference)
                ValidateReference(reference.Value, baseFormats.Length, "命名样式 xfId");
    }

    internal void ValidateSheet(XDocument sheet, string sheetName)
    {
        foreach (var cell in sheet.Descendants(Ns + "c"))
            if (cell.Attribute("s") is { } style)
                ValidateReference(style.Value, _cellStyleCount, $"单元格 {sheetName}!{(string?)cell.Attribute("r")}");
        foreach (var row in sheet.Descendants(Ns + "row"))
            if (row.Attribute("s") is { } style)
                ValidateReference(style.Value, _cellStyleCount, $"工作表 {sheetName} 第 {(string?)row.Attribute("r")} 行");
        foreach (var column in sheet.Descendants(Ns + "col"))
            if (column.Attribute("style") is { } style)
                ValidateReference(style.Value, _cellStyleCount, $"工作表 {sheetName} 第 {(string?)column.Attribute("min")} 列");
    }

    private void ValidateReference(string raw, int count, string location)
    {
        var index = Parse(raw, location);
        if (index >= count)
            throw new InvalidDataException($"{_workbookPath}：{location} 的样式引用 {index} 超出范围 0..{count - 1}。");
    }

    private int Parse(string? raw, string location)
    {
        if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var index) || index < 0)
            throw new InvalidDataException($"{_workbookPath}：{location} 的样式引用无效：{raw}。");
        return index;
    }
}
