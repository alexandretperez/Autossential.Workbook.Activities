namespace Autossential.Workbook.Activities.Core
{
    internal readonly record struct CellRef(int Col, int Row)
    {
        public static bool operator >(CellRef left, CellRef right) => left.Col > right.Col || (left.Col == right.Col && left.Row > right.Row);
        public static bool operator <(CellRef left, CellRef right) => left.Col < right.Col || (left.Col == right.Col && left.Row < right.Row);
        public static bool operator >=(CellRef left, CellRef right) => left > right || left == right;
        public static bool operator <=(CellRef left, CellRef right) => left < right || left == right;
        public static CellRef Parse(ReadOnlySpan<char> span)
        {
            int i = 0;

            if (span[i] == '$')
                i++;

            int col = 0;

            while (i < span.Length)
            {
                var c = span[i];

                if (c >= 'a' && c <= 'z') 
                    c = (char)(c - 32);

                if (c < 'A' || c > 'Z') 
                    break;

                col = (col * 26) + (c - 'A' + 1);
                i++;
            }

            if (col == 0)
                return new CellRef(0, 0);

            if (i < span.Length && span[i] == '$')
                i++;

            int row = 0;

            while (i < span.Length)
            {
                var c = span[i];

                if (c < '0' || c > '9')
                    return new CellRef(0, 0);

                row = (row * 10) + (c - '0');
                i++;
            }

            return new CellRef(col, row);
        }

        public static string GetColumnName(int col)
        {
            if (col <= 0)
                return string.Empty;

            Span<char> buffer = stackalloc char[7];
            var index = buffer.Length;

            while (col > 0)
            {
                col--;
                buffer[--index] = (char)('A' + col % 26);
                col /= 26;
            }

            return new string(buffer[index..]);
        }

        public static int GetColumnIndex(ReadOnlySpan<char> columnName)
        {
            if (columnName.IsEmpty)
                throw new FormatException("Column name cannot be empty.");

            int result = 0;

            foreach (var c in columnName)
            {
                int digit;
                if (c is >= 'A' and <= 'Z')
                    digit = c - 'A' + 1;
                else if (c is >= 'a' and <= 'z')
                    digit = c - 'a' + 1;
                else
                    throw new FormatException($"Invalid column character '{c}' in '{columnName}'.");

                result = result * 26 + digit;

                if (result > 16384) // XFD
                    throw new FormatException($"Column '{columnName}' exceeds maximum column index.");
            }

            return result;
        }

        public static CellRef MaxOpenXML() => new(16_384, 1_048_576);
        public static CellRef MaxBIFF8() => new(256, 65_536);

        public string GetAddress() => $"{GetColumnName(Col)}{Row}";
    }
}