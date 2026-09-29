using System.IO.Compression;
using System.Xml.Linq;

namespace LubanExcelMerge.OpenXml;

internal sealed class OpenXmlStyleImporter
{
    private const int FirstCustomNumberFormatId = 164;
    private readonly XDocument? _target;
    private readonly XDocument? _targetTheme;
    private readonly Dictionary<XDocument, XDocument?> _sourceThemes = new();
    private readonly Dictionary<string, XDocument> _sources = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<(string Path, int StyleIndex), int> _styleMappings = new();

    internal OpenXmlStyleImporter(ZipArchive targetArchive)
    {
        _targetTheme = ReadTheme(targetArchive);
        _target = targetArchive.GetEntry("xl/styles.xml") is null
            ? null
            : OpenXmlWorkbookReader.LoadXml(targetArchive, "xl/styles.xml");
    }

    internal bool IsModified { get; private set; }
    internal XDocument TargetDocument => _target
        ?? throw new InvalidOperationException("目标工作簿没有可写入的样式表。");

    internal string? Resolve(string? styleIndex, string? sourceWorkbook)
    {
        if (styleIndex is null)
            return null;
        if (!int.TryParse(styleIndex, out var parsed) || parsed < 0)
            throw new InvalidDataException($"无效的单元格样式编号：{styleIndex}。");
        if (_target is null)
        {
            if (parsed == 0)
            {
                if (string.IsNullOrWhiteSpace(sourceWorkbook))
                    return null;
                using var sourceArchive = ZipFile.OpenRead(sourceWorkbook);
                if (sourceArchive.GetEntry("xl/styles.xml") is null)
                    return null;
            }
            throw new InvalidDataException("目标工作簿缺少 styles.xml，无法安全导入来源样式。");
        }
        if (string.IsNullOrWhiteSpace(sourceWorkbook))
        {
            ValidateStyleIndex(_target, parsed, "目标工作簿");
            return styleIndex;
        }

        var sourcePath = Path.GetFullPath(sourceWorkbook);
        var key = (sourcePath.ToUpperInvariant(), parsed);
        if (_styleMappings.TryGetValue(key, out var mapped))
            return mapped.ToString();
        var source = GetSource(sourcePath);
        ValidateStyleIndex(source, parsed, sourcePath);
        mapped = MapCellFormat(source, parsed, "cellXfs");
        _styleMappings[key] = mapped;
        return mapped.ToString();
    }

    private XDocument GetSource(string path)
    {
        if (_sources.TryGetValue(path, out var document))
            return document;
        using var archive = ZipFile.OpenRead(path);
        document = OpenXmlWorkbookReader.LoadXml(archive, "xl/styles.xml");
        _sourceThemes[document] = ReadTheme(archive);
        _sources[path] = document;
        return document;
    }

    private int MapCellFormat(XDocument source, int index, string collectionName)
    {
        var sourceCollection = GetCollection(source, collectionName);
        if (index >= sourceCollection.Elements().Count())
            throw new InvalidDataException($"来源工作簿的 {collectionName} 不包含编号 {index}。");
        var clone = new XElement(sourceCollection.Elements().ElementAt(index));
        RemapComponentIndex(source, clone, "fontId", "fonts");
        RemapComponentIndex(source, clone, "fillId", "fills");
        RemapComponentIndex(source, clone, "borderId", "borders");
        RemapNumberFormat(source, clone);
        if (collectionName == "cellXfs" && int.TryParse((string?)clone.Attribute("xfId"), out var baseStyle))
            clone.SetAttributeValue("xfId", MapCellFormat(source, baseStyle, "cellStyleXfs"));
        return FindOrAppend(GetCollection(TargetDocument, collectionName), clone);
    }

    private void RemapComponentIndex(XDocument source, XElement format, string attributeName, string collectionName)
    {
        if (!int.TryParse((string?)format.Attribute(attributeName), out var sourceIndex))
            return;
        var sourceCollection = GetCollection(source, collectionName);
        if (sourceIndex < 0 || sourceIndex >= sourceCollection.Elements().Count())
            throw new InvalidDataException($"来源样式的 {attributeName}={sourceIndex} 超出 {collectionName} 范围。");
        var component = new XElement(sourceCollection.Elements().ElementAt(sourceIndex));
        ValidateComponentContext(source, component);
        format.SetAttributeValue(attributeName, FindOrAppend(GetCollection(TargetDocument, collectionName), component));
    }

    private static XDocument? ReadTheme(ZipArchive archive)
    {
        var relationships = OpenXmlWorkbookReader.LoadXml(archive, "xl/_rels/workbook.xml.rels");
        var relationship = relationships.Root?.Elements().FirstOrDefault(element =>
            ((string?)element.Attribute("Type"))?.EndsWith("/theme", StringComparison.Ordinal) == true);
        if (relationship is null)
            return null;
        var part = OpenXmlWorkbookReader.ResolvePartPath("xl/workbook.xml", (string)relationship.Attribute("Target")!);
        using var stream = (archive.GetEntry(part) ?? throw new InvalidDataException($"工作簿缺少主题 {part}。")).Open();
        return XDocument.Load(stream);
    }

    private void ValidateComponentContext(XDocument source, XElement component)
    {
        var nodes = component.DescendantsAndSelf().ToArray();
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement? ThemePart(XDocument? theme, string name) =>
            theme?.Root?.Element(drawing + "themeElements")?.Element(drawing + name);
        if (nodes.Any(node => node.Attribute("theme") is not null) &&
            !XNode.DeepEquals(ThemePart(_sourceThemes[source], "clrScheme"), ThemePart(_targetTheme, "clrScheme")))
            throw new InvalidDataException("来源与目标工作簿的主题颜色不同，无法安全导入依赖主题的样式，请先统一主题颜色。");
        if (nodes.Any(node => node.Name == OpenXmlNamespaces.Spreadsheet + "scheme") &&
            !XNode.DeepEquals(ThemePart(_sourceThemes[source], "fontScheme"), ThemePart(_targetTheme, "fontScheme")))
            throw new InvalidDataException("来源与目标工作簿的主题字体不同，无法安全导入依赖主题的样式，请先统一主题字体。");
        if (nodes.Any(node => int.TryParse((string?)node.Attribute("indexed"), out var index) && index < 64) &&
            !XNode.DeepEquals(source.Root?.Element(OpenXmlNamespaces.Spreadsheet + "colors"),
                TargetDocument.Root?.Element(OpenXmlNamespaces.Spreadsheet + "colors")))
            throw new InvalidDataException("来源与目标工作簿的索引调色板不同，无法安全导入依赖调色板的样式。");
    }
    private void RemapNumberFormat(XDocument source, XElement format)
    {
        if (!int.TryParse((string?)format.Attribute("numFmtId"), out var sourceId) || sourceId < FirstCustomNumberFormatId)
            return;
        var sourceFormats = source.Root?.Element(OpenXmlNamespaces.Spreadsheet + "numFmts");
        var sourceFormat = sourceFormats?.Elements(OpenXmlNamespaces.Spreadsheet + "numFmt")
            .FirstOrDefault(item => (string?)item.Attribute("numFmtId") == sourceId.ToString())
            ?? throw new InvalidDataException($"来源样式引用了不存在的自定义数字格式 {sourceId}。");
        var formatCode = (string?)sourceFormat.Attribute("formatCode") ?? string.Empty;
        var targetFormats = GetOrCreateCollection(TargetDocument, "numFmts", "fonts");
        var existing = targetFormats.Elements(OpenXmlNamespaces.Spreadsheet + "numFmt")
            .FirstOrDefault(item => string.Equals((string?)item.Attribute("formatCode"), formatCode, StringComparison.Ordinal));
        if (existing is not null)
        {
            format.SetAttributeValue("numFmtId", (string?)existing.Attribute("numFmtId"));
            return;
        }

        var usedIds = targetFormats.Elements(OpenXmlNamespaces.Spreadsheet + "numFmt")
            .Select(item => int.TryParse((string?)item.Attribute("numFmtId"), out var value) ? value : 0)
            .ToHashSet();
        var targetId = FirstCustomNumberFormatId;
        while (usedIds.Contains(targetId))
            targetId++;
        var clone = new XElement(sourceFormat);
        clone.SetAttributeValue("numFmtId", targetId);
        targetFormats.Add(clone);
        UpdateCount(targetFormats);
        IsModified = true;
        format.SetAttributeValue("numFmtId", targetId);
    }

    private int FindOrAppend(XElement targetCollection, XElement candidate)
    {
        var items = targetCollection.Elements().ToArray();
        for (var index = 0; index < items.Length; index++)
        {
            if (XNode.DeepEquals(items[index], candidate))
                return index;
        }
        targetCollection.Add(candidate);
        UpdateCount(targetCollection);
        IsModified = true;
        return items.Length;
    }

    private static XElement GetCollection(XDocument document, string name) =>
        document.Root?.Element(OpenXmlNamespaces.Spreadsheet + name)
        ?? throw new InvalidDataException($"工作簿样式表缺少 {name}。");

    private static XElement GetOrCreateCollection(XDocument document, string name, string insertBefore)
    {
        var existing = document.Root?.Element(OpenXmlNamespaces.Spreadsheet + name);
        if (existing is not null)
            return existing;
        var created = new XElement(OpenXmlNamespaces.Spreadsheet + name, new XAttribute("count", "0"));
        var next = document.Root?.Element(OpenXmlNamespaces.Spreadsheet + insertBefore);
        if (next is null)
            document.Root?.AddFirst(created);
        else
            next.AddBeforeSelf(created);
        return created;
    }

    private static void UpdateCount(XElement collection) =>
        collection.SetAttributeValue("count", collection.Elements().Count());

    private static void ValidateStyleIndex(XDocument styles, int index, string description)
    {
        var count = GetCollection(styles, "cellXfs").Elements().Count();
        if (index >= count)
            throw new InvalidDataException($"{description}的单元格样式编号 {index} 超出 cellXfs 范围 0..{count - 1}。");
    }
}
