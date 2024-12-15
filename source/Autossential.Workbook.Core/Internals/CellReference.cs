using System;
using System.Text;

namespace Autossential.Workbook.Core.Internals
{
    internal abstract class CellReference
    {
        public uint Row { get; set; }
        public uint Col { get; set; }

        protected CellReference(uint row, uint col)
        {
            Row = row;
            Col = col;
            ValidateConstraints();
        }

        protected CellReference(string address)
        {
            ParseAddress(address);
            ValidateConstraints();
        }

        public bool IsValid => Row > 0 && Col > 0;
        protected void ParseAddress(string address)
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

        public string Inspect() => $"{ToString()} (Col: {Col}; Row: {Row}; IsValid: {IsValid})";

        public string GetColumnName()
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

        public override string ToString()
        {
            if (!IsValid) return string.Empty;
            return $"{GetColumnName()}{Row}";
        }

        protected abstract void ValidateConstraints();
    }
}
