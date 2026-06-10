namespace Autossential.Workbook.Activities.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Data;

    public static class TableUtils
    {
        private static readonly string[] SampleNames = { "Ana", "Bruno", "Carlos", "Daniela", "Eduardo", "Fernanda", "Gabriel", "Helena" };
        private static readonly string[] SampleDomains = { "gmail.com", "hotmail.com", "outlook.com", "empresa.com" };
        private static readonly string[] SampleCompanies = { "TechCorp", "Inova Ltda", "Global Solutions", "AlphaSoft", "Beta Systems", "Delta Consulting" };
        private static readonly string[] SampleStreets = { "Rua das Flores", "Avenida Brasil", "Rua XV de Novembro", "Rua das Palmeiras", "Avenida Paulista" };

        private static readonly List<Func<Random, object>> Generators =
        [
            rand => SampleNames[rand.Next(SampleNames.Length)], // Nome
            rand => $"{SampleNames[rand.Next(SampleNames.Length)].ToLower()}{rand.Next(100)}@{SampleDomains[rand.Next(SampleDomains.Length)]}", // Email
            rand => DateTime.Today.AddDays(rand.Next(-1000, 1000)).ToString("yyyy-MM-dd"), // Data
            rand => new TimeSpan(rand.Next(0, 24), rand.Next(0, 60), 0).ToString(@"hh\:mm"), // Hora
            rand => DateTime.Now.AddMinutes(rand.Next(-10000, 10000)).ToString("yyyy-MM-dd HH:mm"), // Data e Hora
            rand => rand.Next(2) == 0 ? "true" : "false", // Booleano
            rand => rand.Next(-1000, 1000), // Inteiro
            rand => (rand.NextDouble() * 1000).ToString("F2"), // Double
            rand => DBNull.Value, // Null/vazio
            rand => SampleCompanies[rand.Next(SampleCompanies.Length)], // Companhia
            rand => $"{SampleStreets[rand.Next(SampleStreets.Length)]}, {rand.Next(1, 999)}", // Endereço
        ];

        public static DataTable Generate(int cols, int rows, string seed)
        {
            var table = new DataTable();
            int hashSeed = seed.GetHashCode() ^ cols ^ rows;
            Random rand = new(hashSeed);

            // Criar colunas genéricas
            for (int c = 0; c < cols; c++)
            {
                table.Columns.Add($"Col{c + 1}", typeof(object));
            }

            // Popular linhas
            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    var generator = Generators[rand.Next(Generators.Count)];
                    row[c] = generator(rand);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        public static DataTable Build(int cols, int rows, Func<int, int, object?> value)
        {
            var table = new DataTable();

            // Criar colunas genéricas
            for (int c = 0; c < cols; c++)
            {
                table.Columns.Add($"Col{c + 1}", typeof(object));
            }

            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    row[c] = value(c + 1, r + 1);
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
