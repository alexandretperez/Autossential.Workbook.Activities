namespace Autossential.Workbook.Activities.Core
{
    internal abstract record class CellReference
    {
        private void Initialize(int col, int row)
        {
            var (maxCol, maxRow) = MaxReference();
            if (col < 1 || col > maxCol)
                throw new ArgumentOutOfRangeException(nameof(col), "The 'col' is out of range for this workbook type.");

            if (row < 1 || row > maxRow)
                throw new ArgumentOutOfRangeException(nameof(row), "The 'row' is out of range for this workbook type.");

            Col = col;
            Row = row;
        }

        protected CellReference(int col, int row) => Initialize(col, row);
        protected CellReference(string address, int maxRows)
        {
            var (col, row, inferredRow) = Parse(address, maxRows);
            Initialize(col, row);
            _isRowInferred = inferredRow;
        }

        public int Col { get; private set; }
        public int Row { get; private set; }

        private readonly bool _isRowInferred;
        public bool IsRowInferred() => _isRowInferred;
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
        public string ToAddress() => $"{GetColumnName(Col)}{Row}";
        public static (int col, int row, bool inferredRow) Parse(string address, int maxRows)
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
                return (col, maxRows, true);

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

        public static bool operator >(CellReference left, CellReference right) => left.Col > right.Col || (left.Col == right.Col && left.Row > right.Row);
        public static bool operator <(CellReference left, CellReference right) => left.Col < right.Col || (left.Col == right.Col && left.Row < right.Row);
        public static bool operator >=(CellReference left, CellReference right) => left > right || left == right;
        public static bool operator <=(CellReference left, CellReference right) => left < right || left == right;
        public override int GetHashCode() => HashCode.Combine(Col, Row);
        public virtual bool Equals(CellReference other) => other is not null && Col == other.Col && Row == other.Row;
        protected abstract (int col, int row) MaxReference();
    }
}