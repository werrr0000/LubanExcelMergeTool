using System.Globalization;

namespace LubanExcelMerge.Luban;

public static class LubanKeyValueNormalizer
{
    public static string? Normalize(LubanField field, string? value)
    {
        if (IsBoolean(field.TypeName))
        {
            if (string.IsNullOrEmpty(value))
                return "0";
            if (value == "1" || bool.TryParse(value, out var boolean) && boolean)
                return "1";
            if (value == "0" || bool.TryParse(value, out boolean) && !boolean)
                return "0";
            return value;
        }

        if (!IsNumeric(field.TypeName))
            return value;
        if (string.IsNullOrEmpty(value))
            return "0";
        return decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number.ToString("G29", CultureInfo.InvariantCulture)
            : value;
    }

    private static bool IsBoolean(string typeName) =>
        string.Equals(typeName, "bool", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(typeName, "boolean", StringComparison.OrdinalIgnoreCase);

    private static bool IsNumeric(string typeName)
    {
        var simpleName = typeName.Split('.').LastOrDefault() ?? typeName;
        return simpleName.Equals("FPString", StringComparison.OrdinalIgnoreCase) ||
               simpleName is "byte" or "sbyte" or "short" or "ushort" or "int" or "uint" or
                   "long" or "ulong" or "float" or "double" or "decimal";
    }
}
