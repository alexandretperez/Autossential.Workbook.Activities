using System.Text;

namespace Autossential.Workbook.Activities.Core
{
    internal class CellReference : IEquatable<CellReference>
    {
        public const int BIFF8_MAX_COLS = 256;
        public const int BIFF8_MAX_ROWS = 65_536;
        public const int OPENXML_MAX_COLS = 16_384;
        public const int OPENXML_MAX_ROWS = 1_048_576;

        public CellReference(ExcelFileType type, string address) : this(type, ParseAddress(address))
        {
        }

        public CellReference(ExcelFileType type, int col, int row) : this(type, (col, row))
        {
        }

        private CellReference(ExcelFileType type, (int col, int row) coordinates)
        {
            var (col, row) = coordinates;
            var outOfRange = type switch
            {
                ExcelFileType.Binary => (col < 1 || col > BIFF8_MAX_COLS || row < 1 || row > BIFF8_MAX_ROWS),
                ExcelFileType.OpenXML => (col < 1 || col > OPENXML_MAX_COLS || row < 1 || row > OPENXML_MAX_ROWS),
                _ => throw new InvalidOperationException("Unsupported Excel file type.")
            };

            if (outOfRange)
                throw new InvalidOperationException("Col or row reference is out of the supported limit for the specified workbook file type.");

            Row = row;
            Col = col;
        }

        public int Col { get; }
        public bool IsValid => Row > 0 && Col > 0;
        public int Row { get; }

        public static CellReference Max(ExcelFileType type)
        {
            return type switch
            {
                ExcelFileType.Binary => new CellReference(ExcelFileType.Binary, BIFF8_MAX_COLS, BIFF8_MAX_ROWS),
                ExcelFileType.OpenXML => new CellReference(ExcelFileType.OpenXML, OPENXML_MAX_COLS, OPENXML_MAX_ROWS),
                _ => throw new InvalidOperationException("Unsupported Excel file type.")
            };
        }

        public string GetColumnName()
        {
            if (Col == 0) return string.Empty;

            var columnName = new StringBuilder();
            int dividend = Col;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName.Insert(0, (char)(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString();
        }

        public override string ToString()
        {
            return $"{GetColumnName()}{Row}";
        }

        private static (int Col, int Row) ParseAddress(string address)
        {
            if (string.IsNullOrEmpty(address))
                return (0, 0);

            int col = 0;
            const int offset = 'A' - 1;

            int i = -1;
            int len = address.Length;
            while (++i < len)
            {
                char c = address[i];
                if (c >= 'A' && c <= 'Z')
                {
                    col *= 26;
                    col += c - offset;
                    continue;
                }
                else if (c == '$' && (i == 0 || char.IsDigit(address[Math.Min(i + 1, len - 1)])))
                {
                    continue;
                }
                break;
            }

            if (col == 0 || !int.TryParse(address.AsSpan(i), out int row))
            {
                return (0, 0);
            }

            return (col, row);
        }

        public bool Equals(CellReference other)
        {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            return Row == other.Row && Col == other.Col;
        }

        public override bool Equals(object obj) => Equals(obj as CellReference);

        public override int GetHashCode() => HashCode.Combine(Row, Col);

        public static bool operator ==(CellReference left, CellReference right)
            => EqualityComparer<CellReference>.Default.Equals(left, right);

        public static bool operator !=(CellReference left, CellReference right)
            => !(left == right);
    }
}