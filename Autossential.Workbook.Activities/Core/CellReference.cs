namespace Autossential.Workbook.Activities.Core
{
    internal readonly struct CellReference : IEquatable<CellReference>
    {
        public const string MIN_REFERENCE = "A1";

        public const int BIFF8_MAX_COLS = 256;
        public const int BIFF8_MAX_ROWS = 65_536;
        public const string BIFF8_MAX_REFERENCE = "IV65536";

        public const int OPENXML_MAX_COLS = 16_384;
        public const int OPENXML_MAX_ROWS = 1_048_576;
        public const string OPENXML_MAX_REFERENCE = "XFD1048576";

        private CellReference(int col, int row)
        {
            Row = row;
            Col = col;
        }

        public static CellReference OpenXml(int col, int row)
        {
            ValidateOpenXml(col, row);
            return new CellReference(col, row);
        }

        public static CellReference OpenXml(string address)
        {
            if (OPENXML_MAX_REFERENCE.Equals(address, StringComparison.OrdinalIgnoreCase))
                return OpenXmlMax();

            var (col, row) = Parse(address);
            ValidateOpenXml(col, row);
            return new CellReference(col, row);
        }

        public static CellReference OpenXmlMax() => new(OPENXML_MAX_COLS, OPENXML_MAX_ROWS);

        public static CellReference Binary(int col, int row)
        {
            ValidateBinary(col, row);
            return new CellReference(col, row);
        }

        public static CellReference Binary(string address)
        {
            if (BIFF8_MAX_REFERENCE.Equals(address, StringComparison.OrdinalIgnoreCase))
                return BinaryMax();

            var (col, row) = Parse(address);
            ValidateBinary(col, row);
            return new CellReference(col, row);
        }

        public static CellReference BinaryMax() => new(BIFF8_MAX_COLS, BIFF8_MAX_ROWS);

        private static void ValidateOpenXml(int col, int row)
        {
            if (col < 1 || col > OPENXML_MAX_COLS || row < 1 || row > OPENXML_MAX_ROWS)
                throw new InvalidOperationException("Row or col is out of the supported limit for Excel (.xlsx files)");
        }

        private static void ValidateBinary(int col, int row)
        {
            if (col < 1 || col > BIFF8_MAX_COLS || row < 1 || row > BIFF8_MAX_ROWS)
                throw new InvalidOperationException("Row or col is out of the supported limit for Excel 97–2003 (.xls files)");
        }

        public int Row { get; }
        public int Col { get; }
        public bool IsValid => Row > 0 && Col > 0;

        private static (int Col, int Row) Parse(string address)
        {
            if (string.IsNullOrEmpty(address))
                return (0, 0);

            var span = address.AsSpan();
            var i = 0;

            if (span[i] == '$')
                i++;

            var col = 0;
            var hasColumn = false;

            while (i < span.Length)
            {
                var c = span[i];

                if (c >= 'a' && c <= 'z')
                    c = (char)(c - 32);

                if (c < 'A' || c > 'Z')
                    break;

                hasColumn = true;
                col = (col * 26) + (c - 'A' + 1);
                i++;
            }

            if (!hasColumn)
                return (0, 0);

            if (i < span.Length && span[i] == '$')
                i++;

            if (i >= span.Length)
                return (0, 0);

            var row = 0;

            while (i < span.Length)
            {
                var c = span[i];

                if (c < '0' || c > '9')
                    return (0, 0);

                row = (row * 10) + (c - '0');
                i++;
            }

            return row > 0 ? (col, row) : (0, 0);
        }

        public readonly bool Equals(CellReference other) =>
            Row == other.Row && Col == other.Col;

        public override readonly bool Equals(object obj) =>
            obj is CellReference other && Equals(other);

        public static bool operator ==(CellReference left, CellReference right) =>
            left.Equals(right);

        public static bool operator !=(CellReference left, CellReference right) =>
            !left.Equals(right);

        public override readonly int GetHashCode() =>
            HashCode.Combine(Col, Row);

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

        public override readonly string ToString() => IsValid ? $"{GetColumnName(Col)}{Row}" : string.Empty;
    }
}