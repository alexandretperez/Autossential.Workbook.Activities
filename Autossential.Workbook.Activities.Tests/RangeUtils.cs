using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Autossential.Workbook.Activities.Tests
{
    public static class RangeUtils
    {
        public static (int colStart, int colEnd, int rowStart, int rowEnd) ParseRange(string range, DataTable dt)
        {
            var parts = range.Split(':');

            // Se só veio uma célula, usamos até o fim da tabela
            if (parts.Length == 1)
            {
                var (col, row) = ParseCell(parts[0]);
                return (col, dt.Columns.Count - 1, row, dt.Rows.Count - 1);
            }
            else if (parts.Length == 2)
            {
                var (colStart, rowStart) = ParseCell(parts[0]);
                var (colEnd, rowEnd) = ParseCell(parts[1]);
                return (colStart, colEnd, rowStart, rowEnd);
            }
            else
            {
                throw new ArgumentException("Range inválido");
            }
        }

        public static (int col, int row) ParseCell(string cell)
        {
            int i = 0;
            while (i < cell.Length && char.IsLetter(cell[i])) i++;

            string colLetters = cell[..i].ToUpper();
            string rowDigits = cell[i..];

            int col = 0;
            foreach (char ch in colLetters)
            {
                col = col * 26 + (ch - 'A' + 1);
            }
            col--; // zero-based

            int row = int.Parse(rowDigits) - 1; // zero-based

            return (col, row);
        }
    }
}
