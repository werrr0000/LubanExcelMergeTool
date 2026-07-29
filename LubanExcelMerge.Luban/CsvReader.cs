using System.Text;

namespace LubanExcelMerge.Luban;

public static class CsvReader
{
    public static IReadOnlyList<IReadOnlyList<string>> ReadAll(TextReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);

        var rows = new List<IReadOnlyList<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        var closedQuote = false;

        while (reader.Read() is var value && value >= 0)
        {
            var character = (char)value;
            if (inQuotes)
            {
                if (character != '"')
                {
                    field.Append(character);
                    continue;
                }

                if (reader.Peek() == '"')
                {
                    reader.Read();
                    field.Append('"');
                    continue;
                }

                inQuotes = false;
                closedQuote = true;
                continue;
            }

            if (character == '"')
            {
                if (field.Length != 0 || closedQuote)
                    throw new FormatException("CSV 引号只能出现在字段开头。");

                inQuotes = true;
                continue;
            }

            if (character == ',')
            {
                AddField(row, field);
                closedQuote = false;
                continue;
            }

            if (character is '\r' or '\n')
            {
                if (character == '\r' && reader.Peek() == '\n')
                    reader.Read();

                AddField(row, field);
                rows.Add(row);
                row = new List<string>();
                closedQuote = false;
                continue;
            }

            if (closedQuote)
                throw new FormatException("CSV 引号字段结束后只能出现分隔符或换行符。");

            field.Append(character);
        }

        if (inQuotes)
            throw new FormatException("CSV 文件包含未闭合的引号字段。");

        if (row.Count > 0 || field.Length > 0 || closedQuote)
        {
            AddField(row, field);
            rows.Add(row);
        }

        return rows;
    }

    private static void AddField(List<string> row, StringBuilder field)
    {
        row.Add(field.ToString());
        field.Clear();
    }
}
