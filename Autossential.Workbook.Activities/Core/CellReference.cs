namespace Autossential.Workbook.Activities.Core
{
    public readonly struct CellReference : IEquatable<CellReference>
    {
        public const int BIFF8_MAX_COLS = 256;
        public const string BIFF8_MAX_REFERENCE = "IV65536";
        public const int BIFF8_MAX_ROWS = 65_536;
        public const int OPENXML_MAX_COLS = 16_384;
        public const string OPENXML_MAX_REFERENCE = "XFD1048576";
        public const int OPENXML_MAX_ROWS = 1_048_576;

        public CellReference(int col, int row)
        {
            if (col < 1) throw new ArgumentOutOfRangeException(nameof(col), "The 'col' argument must be greater than zero.");
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(col), "The 'row' argument must be greater than zero.");

            Col = col;
            Row = row;
        }

        public CellReference(string address)
        {
            var (col, row, inferredRow) = Parse(address);
            Col = col;
            Row = row;
            _hasInferredRow = inferredRow;
        }

        public int Col { get; }
        public bool IsValidForBIFF8() => Col > 0 && Col <= BIFF8_MAX_COLS && Row > 0 && Row <= BIFF8_MAX_ROWS;
        public bool IsValidForOpenXml() => Col > 0 && Col <= OPENXML_MAX_COLS && Row > 0 && Row <= OPENXML_MAX_ROWS;
        public int Row { get; }
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

        public static bool operator !=(CellReference left, CellReference right) => !left.Equals(right);

        public static bool operator ==(CellReference left, CellReference right) => left.Equals(right);

        public static bool operator >(CellReference left, CellReference right) => left.Col > right.Col || (left.Col == right.Col && left.Row > right.Row);
        public static bool operator <(CellReference left, CellReference right) => left.Col < right.Col || (left.Col == right.Col && left.Row < right.Row);

        public static bool operator >=(CellReference left, CellReference right) => left > right || left == right;
        public static bool operator <=(CellReference left, CellReference right) => left < right || left == right;

        public bool Equals(CellReference other) => Row == other.Row && Col == other.Col;
        public override bool Equals(object obj) => obj is CellReference other && Equals(other);

        public override readonly int GetHashCode() =>
             HashCode.Combine(Col, Row);

        public bool HasInferredRow() => _hasInferredRow;

        public override readonly string ToString() => $"{GetColumnName(Col)}{Row}";

        private bool _hasInferredRow { get; }
        private static (int col, int row, bool hasInferredRow) Parse(string address)
        {
            if (string.IsNullOrEmpty(address))
                return (0, 0, false);

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
                return (0, 0, false);

            if (i < span.Length && span[i] == '$')
                i++;

            if (i >= span.Length)
                return (col, 1, true);

            var row = 0;

            while (i < span.Length)
            {
                var c = span[i];

                if (c < '0' || c > '9')
                    return (0, 0, false);

                row = (row * 10) + (c - '0');
                i++;
            }

            return row > 0 ? (col, row, false) : (0, 0, false);
        }
    }
}