using NPOI.SS.UserModel;
using System.Data;

namespace Autossential.Workbook.Activities.Tests
{
    public static class TableGenerator
    {
        private static readonly Random _random = new Random();

        public static DataTable GenerateTable(int cols, int rows, params Type[] types)
        {
            if (types == null || types.Length == 0)
                types = _types;

            DataTable table = new();

            // Gerar colunas com tipos sorteados
            for (int i = 0; i < cols; i++)
            {
                Type chosenType = types[_random.Next(types.Length)];
                table.Columns.Add($"Col{i + 1}", chosenType);
            }

            // Preencher linhas com dados aleatórios
            for (int r = 0; r < rows; r++)
            {
                DataRow row = table.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    Type colType = table.Columns[c].DataType;
                    row[c] = GenerateRandomValue(colType);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        private static readonly string[] _wordPool = [
            "sol", "lua", "estrela", "vento", "mar", "montanha", "floresta", "tempo", "caminho", "sonho",
            "esperança", "vida", "amor", "destino", "coragem", "luz", "noite", "dia", "silêncio", "voz"
        ];

        private static readonly Type[] _types = [
            typeof(string),
            typeof(double),
            typeof(bool),
            typeof(int),
            typeof(DateTime)
        ];

        private static object? GenerateRandomValue(Type type)
        {
            if (type == typeof(string))
            {
                int wordCount = _random.Next(1, 6); // Frases de 1 a 5 palavras
                var words = Enumerable.Range(0, wordCount)
                                      .Select(_ => _wordPool[_random.Next(_wordPool.Length)]);
                string sentence = string.Join(" ", words);
                return char.ToUpper(sentence[0]) + sentence.Substring(1); // Capitaliza a primeira letra
            }

            if (type == typeof(double))
                return Math.Round(_random.NextDouble() * 1000, 2);

            if (type == typeof(bool))
                return _random.Next(2) == 0;

            if (type == typeof(int))
                return _random.Next(0, 1000);

            if (type == typeof(DateTime))
                return DateTime.Now.AddDays(_random.Next(-1000, 1000));

            // Valor padrão para tipos não suportados
            return null;
        }
    }
}
