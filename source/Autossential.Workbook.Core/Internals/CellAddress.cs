using System;

namespace Autossential.Workbook.Core.Internals
{
    internal class CellAddress
    {
        public CellAddress(string address)
        {
            if (address == null)
                return;

            const int offset = 'A' - 1;

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

            if (Col == 0 || !int.TryParse(address.Substring(i), out int row))
            {
                Row = 0;
                Col = 0;
                return;
            }

            Row = row;
        }

        public int Row { get; private set; }
        public int Col { get; private set; }
        public bool IsValid => Row * Col > 0;
        public override string ToString() => $"Col: {Col}; Row: {Row}; IsValid: {IsValid}";

        public void SetDefault(int col, int row)
        {
            if ((IsValid && !IsDefault))
                return;

            Override(col, row);
            IsDefault = true;
        }
        public void Override(int col, int row)
        {
            Col = col;
            Row = row;
        }

        public static CellAddress UseDefault(string address)
        {
            return new CellAddress(address) { IsDefault = true };
        }

        public bool IsDefault { get; private set; }
    }
}
