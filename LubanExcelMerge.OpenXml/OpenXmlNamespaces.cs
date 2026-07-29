using System.Xml.Linq;

namespace LubanExcelMerge.OpenXml;

internal static class OpenXmlNamespaces
{
    internal static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static readonly XNamespace OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";
}
