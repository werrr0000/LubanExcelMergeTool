using System.Text;
using System.Text.RegularExpressions;

namespace LubanExcelMerge.OpenXml;

internal static partial class FormulaReferenceShifter
{
    [GeneratedRegex(@"(?<![A-Za-z0-9_.])(?<column>\$?[A-Za-z]{1,3})(?<row>\$?[1-9][0-9]*)(?![A-Za-z0-9_!])")]
    private static partial Regex CellReferencePattern();

    internal static string ShiftRows(string formula, int firstMovedRow, int delta)
    {
        if (delta == 0 || formula.Length == 0)
            return formula;

        var result = new StringBuilder(formula.Length);
        var segmentStart = 0;
        var insideString = false;
        for (var index = 0; index < formula.Length; index++)
        {
            if (formula[index] != '"')
                continue;
            if (insideString && index + 1 < formula.Length && formula[index + 1] == '"')
            {
                index++;
                continue;
            }

            if (!insideString)
                AppendShifted(result, formula.AsSpan(segmentStart, index - segmentStart), firstMovedRow, delta);
            else
                result.Append(formula.AsSpan(segmentStart, index - segmentStart + 1));
            insideString = !insideString;
            segmentStart = index + 1;
        }

        if (insideString)
            result.Append(formula.AsSpan(segmentStart));
        else
            AppendShifted(result, formula.AsSpan(segmentStart), firstMovedRow, delta);
        return result.ToString();
    }

    private static void AppendShifted(StringBuilder result, ReadOnlySpan<char> segment, int firstMovedRow, int delta)
    {
        var text = segment.ToString();
        result.Append(CellReferencePattern().Replace(text, match =>
        {
            var rowText = match.Groups["row"].Value;
            var absolute = rowText[0] == '$';
            var rowNumber = int.Parse(absolute ? rowText[1..] : rowText);
            if (rowNumber < firstMovedRow)
                return match.Value;
            var shifted = rowNumber + delta;
            if (shifted < 1)
                return "#REF!";
            return match.Groups["column"].Value + (absolute ? "$" : string.Empty) + shifted;
        }));
    }
}
