namespace LubanExcelMerge.OpenXml;

public static class CellReference
{
    public static (int RowNumber, int ColumnIndex) Parse(string address)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(address);
        var separator = 0;
        while (separator < address.Length && char.IsLetter(address[separator]))
            separator++;

        if (separator == 0 || separator == address.Length ||
            !int.TryParse(address[separator..], out var rowNumber) || rowNumber < 1)
            throw new FormatException($"无效的单元格地址：{address}。");

        var columnIndex = 0;
        foreach (var character in address[..separator])
        {
            var upper = char.ToUpperInvariant(character);
            if (upper is < 'A' or > 'Z')
                throw new FormatException($"无效的单元格地址：{address}。");
            columnIndex = checked(columnIndex * 26 + upper - 'A' + 1);
        }

        return (rowNumber, columnIndex - 1);
    }

    public static string Create(int rowNumber, int columnIndex)
    {
        if (rowNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));
        if (columnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));

        var letters = string.Empty;
        for (var value = columnIndex + 1; value > 0; value = (value - 1) / 26)
            letters = (char)('A' + (value - 1) % 26) + letters;
        return letters + rowNumber;
    }
}
