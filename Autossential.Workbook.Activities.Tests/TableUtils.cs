using System.Data;

namespace Autossential.Workbook.Activities.Tests
{
    public static class TableUtils
    {
        private static readonly Func<Random, int, int, object?>[] Generators = [
            (rand, col, row) => $"C{col}R{row}",
            (rand, col, row) => DateTime.Today.AddDays(rand.Next(-1000, 1000)).ToString("yyyy-MM-dd"),
            (rand, col, row) => new TimeSpan(rand.Next(0, 24), rand.Next(0, 60), 0).ToString(@"hh\:mm"),
            (rand, col, row) => DateTime.Today.AddMinutes(rand.Next(-10000, 10000)).ToString("yyyy-MM-dd HH:mm"),
            (rand, col, row) => null,
            (rand, col, row) => rand.Next(2) == 0 ,
            (rand, col, row) => rand.Next(-1000, 1000),
            (rand, col, row) => Math.Round(rand.NextDouble() * 1000, 2),
            (rand, col, row) => DBNull.Value,
            (rand, col, row) => string.Empty
        ];

        public static DataTable Build(int cols, int rows, Func<int, int, object?>? valueResolver = null)
        {
            valueResolver ??= (col, row) => $"C{col}R{row}";

            var table = new DataTable();
            if (cols == 0 && rows == 0)
                return table;

            for (int c = 0; c < cols; c++)
                table.Columns.Add($"Col{c + 1}", typeof(object));

            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                    row[c] = valueResolver(c + 1, r + 1);

                table.Rows.Add(row);
            }
            return table;
        }

        public static DataTable Generate(int cols, int rows, int seed = 0)
        {
            int hashSeed = cols ^ rows ^ seed;
            var rnd = new Random(hashSeed);
            var table = new DataTable();
            for (int c = 0; c < cols; c++)
                table.Columns.Add($"Col{c + 1}", typeof(object));

            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    var generator = Generators[rnd.Next(Generators.Length)];
                    row[c] = generator(rnd, c + 1, r + 1);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        public static bool AreTablesEqual(DataTable t1, DataTable t2)
        {
            // Verifica dimensões
            if (t1.Columns.Count != t2.Columns.Count || t1.Rows.Count != t2.Rows.Count)
                return false;

            // Verifica nomes das colunas
            for (int c = 0; c < t1.Columns.Count; c++)
            {
                if (t1.Columns[c].ColumnName != t2.Columns[c].ColumnName)
                    return false;
            }

            static bool LooseEquals(object v1, object v2)
            {
                if (v1 == null && v2 == null) return true;
                if (v1 == null || v2 == null) return false;

                // Números
                if (decimal.TryParse(v1.ToString(), out var d1) &&
                    decimal.TryParse(v2.ToString(), out var d2))
                {
                    return d1 == d2;
                }

                // Booleanos
                if (bool.TryParse(v1.ToString(), out var b1) &&
                    bool.TryParse(v2.ToString(), out var b2))
                {
                    return b1 == b2;
                }

                // Datas
                if (DateTime.TryParse(v1.ToString(), out var dt1) &&
                    DateTime.TryParse(v2.ToString(), out var dt2))
                {
                    return dt1 == dt2;
                }

                // Fallback: comparar como string
                return v1.ToString() == v2.ToString();
            }

            // Verifica valores célula a célula
            for (int r = 0; r < t1.Rows.Count; r++)
            {
                for (int c = 0; c < t1.Columns.Count; c++)
                {
                    var v1 = t1.Rows[r][c];
                    var v2 = t2.Rows[r][c];

                    if (!LooseEquals(v1, v2))
                        return false;
                }
            }

            return true;
        }
    }
}
