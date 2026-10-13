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
        if (!double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out var leftDouble) ||
            !double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out var rightDouble) ||
            !double.IsFinite(leftDouble) || !double.IsFinite(rightDouble))
            return false;

        if (decimal.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out var leftNumber) &&
            decimal.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out var rightNumber))
        {
            // Decimal can underflow valid Excel numbers to zero. Only accept
            // that equality if their double values also match.
            if (leftNumber == rightNumber && (leftNumber != 0 || leftDouble == rightDouble))
                return true;
            // Preserve exact integer distinctions, including IDs beyond 15 digits.
            if (leftNumber != rightNumber && decimal.Truncate(leftNumber) == leftNumber && decimal.Truncate(rightNumber) == rightNumber)
                return false;
        }
        // Excel numeric cells compare at 15 significant digits. This removes
        // serializer tails without rounding or rewriting any source payload.
        return leftDouble.ToString("G15", CultureInfo.InvariantCulture) == rightDouble.ToString("G15", CultureInfo.InvariantCulture);
    }
}
