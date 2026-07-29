using System.Globalization;

namespace LubanExcelMerge.Core;

public enum CellValueKind
{
    Blank,
    Boolean,
    Number,
    String,
    Error,
    Formula
}

public sealed record CellPayload(
    CellValueKind Kind,
    string? RawValue = null,
    string? FormulaText = null,
    string? CachedValue = null,
    string? SourceWorkbook = null,
    string? SourceSheet = null,
    string? Address = null,
    string? RawDataType = null,
    bool HasExternalWorkbookReference = false,
    IReadOnlyDictionary<string, string>? FormulaAttributes = null)
{
    public static CellPayload Blank { get; } = new(CellValueKind.Blank);

    public bool ContentEquals(CellPayload? other)
    {
        if (other is null || Kind != other.Kind)
            return false;

        return Kind switch
        {
            CellValueKind.Blank => true,
            CellValueKind.Number => NumbersEqual(RawValue, other.RawValue),
            CellValueKind.Formula => string.Equals(FormulaText, other.FormulaText, StringComparison.Ordinal),
            _ => string.Equals(RawValue, other.RawValue, StringComparison.Ordinal)
        };
    }

    private static bool NumbersEqual(string? left, string? right)
    {
        return decimal.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out var leftNumber)
            && decimal.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out var rightNumber)
            && leftNumber == rightNumber;
    }
}
