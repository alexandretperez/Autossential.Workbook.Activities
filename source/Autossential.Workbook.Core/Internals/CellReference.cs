using NPOI.SS.Formula.Functions;
using System;
using System.Text;

namespace Autossential.Workbook.Core.Internals
{
    internal struct CellReference
    {
        private const uint BIFF8_MAX_COLS = 256;
        private const uint BIFF8_MAX_ROWS = 65_536;
        private const uint OPENXML_MAX_COLS = 16_384;
        private const uint OPENXML_MAX_ROWS = 1_048_576;

        public CellReference() // A1 by default
        {
            Row = 1;
            Col = 1;
        }

        private CellReference(uint row, uint col)
        {
            Row = row;
            Col = col;
        }

        private CellReference(string address)
        {
            Row = 0;
            Col = 0;
            ParseAddress(address);
        }

        public uint Col { get; set; }

        public readonly bool IsValid => Row > 0 && Col > 0;

        public uint Row { get; set; }

        public static CellReference CreateForBIFF8(uint row, uint col)
        {
            var cellRef = new CellReference(row, col);
            cellRef.ValidateBIFF8Constraints();
            return cellRef;
        }
        public static CellReference CreateForBIFF8(string cellAddress)
        {
            var cellRef = new CellReference(cellAddress);
            cellRef.ValidateBIFF8Constraints();
            return cellRef;
        }
        public static CellReference CreateForOpenXml(uint row, uint col)
        {
            var cellRef = new CellReference(row, col);
            cellRef.ValidateOpenXMLConstraints();
            return cellRef;
        }
        public static CellReference CreateForOpenXml(string cellAddress)
        {
            var cellRef = new CellReference(cellAddress);
            cellRef.ValidateOpenXMLConstraints();
            return cellRef;
        }
        public static CellReference GetMaxForBIFF8() => new(BIFF8_MAX_ROWS, BIFF8_MAX_COLS);
        public static CellReference GetMaxForOpenXml() => new(OPENXML_MAX_ROWS, OPENXML_MAX_COLS);

        public readonly string GetColumnName()
        {
            if (Col == 0) return string.Empty;

            var columnName = new StringBuilder();
            uint dividend = Col;
            uint modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName.Insert(0, (char)(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString();
        }

        public void ParseAddress(string address)
        {
            if (string.IsNullOrEmpty(address))
                return;

            const uint offset = 'A' - 1;

            int i = -1;
            int len = address.Length;
            while (++i < len)
            {
                char c = address[i];
                if (c >= 'A' && c <= 'Z')
                {
                    Col *= 26;
                    Col += c - offset;
                    continue;
                }
                else if (c == '$' && (i == 0 || char.IsDigit(address[Math.Min(i + 1, len - 1)])))
                {
                    continue;
                }
                break;
            }

            if (Col == 0 || !uint.TryParse(address.AsSpan(i), out uint row))
            {
                Row = 0;
                Col = 0;
                return;
            }

            Row = row;
        }

        public override readonly string ToString()
        {
            return $"{GetColumnName()}{Row}";
        }

        private void ValidateBIFF8Constraints()
        {
            if (Row > BIFF8_MAX_ROWS || Col > BIFF8_MAX_COLS)
                Row = Col = 0;
        }

        private void ValidateOpenXMLConstraints()
        {
            if (Row > OPENXML_MAX_ROWS || Col > OPENXML_MAX_COLS)
                Row = Col = 0;
        }

        public readonly bool IsRowInRange(long start, long end) => Row >= start && Row <= end;
        public readonly bool IsColInRange(long start, long end) => Col >= start && Col <= end;
    }
}