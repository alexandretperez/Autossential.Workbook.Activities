using Autossential.Workbook.Activities.Extensions;
using ExcelDataReader;
using System.Data;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal abstract class WorkbookProcessorBase : IWorkbookProcessor
    {
        IExcelDataReader _reader;

        protected WorkbookProcessorBase(string filePath, string password)
        {
            FilePath = filePath;
            Password = password;
            WorkbookStream = new MemoryStream();
            if (File.Exists(FilePath))
            {
                var bytes = File.ReadAllBytes(FilePath);
                WorkbookStream.Write(bytes, 0, bytes.Length);
                WorkbookStream.Reset();
            }
        }

        public string FilePath { get; }
        public string Password { get; }
        public MemoryStream WorkbookStream { get; }

        public void Dispose()
        {
            _reader?.Dispose();
            WorkbookStream?.Dispose();
        }

        public int GetColumnCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range);
            var reader = GetReader();

            do
            {
                if (reader.Name != sheetName)
                    continue;

                var row = 0;
                var firstCol = int.MaxValue;
                var lastCol = int.MinValue;

                while (reader.Read())
                {
                    ++row;
                    if (row < rangeRef.Start.Row) continue;
                    if (row > rangeRef.End.Row) break;

                    var size = Math.Min(rangeRef.End.Col, reader.FieldCount);

                    for (int col = rangeRef.Start.Col; col <= size; col++)
                    {
                        var value = reader.GetValue(col - 1);
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            continue;

                        if (col < firstCol) firstCol = col;
                        if (col > lastCol) lastCol = col;
                    }

                    // Early exit: range inteira preenchida
                    if (firstCol == rangeRef.Start.Col && lastCol == size)
                        return size - rangeRef.Start.Col + 1;
                }

                return firstCol == int.MaxValue ? 0 : lastCol - firstCol + 1;

            } while (reader.NextResult());

            return 0;
        }

        public int GetRowCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range);
            var reader = GetReader();

            do
            {
                if (reader.Name != sheetName)
                    continue;

                var row = 0;
                var firstRow = int.MaxValue;
                var lastRow = int.MinValue;

                while (reader.Read())
                {
                    ++row;
                    if (row < rangeRef.Start.Row) continue;
                    if (row > rangeRef.End.Row) break;

                    var maxCol = Math.Min(rangeRef.End.Col, reader.FieldCount);
                    for (int col = rangeRef.Start.Col; col <= maxCol; col++)
                    {
                        var value = reader.GetValue(col - 1);
                        if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                            continue;

                        if (row < firstRow) firstRow = row;
                        if (row > lastRow) lastRow = row;
                        break;
                    }
                }

                return firstRow == int.MaxValue ? 0 : lastRow - firstRow + 1;

            } while (reader.NextResult());

            return 0;
        }

        public string[] GetSheetNames()
        {
            var reader = GetReader();
            var sheetNames = new string[reader.ResultsCount];
            int i = 0;
            do
            {
                sheetNames[i++] = reader.Name;
            } while (reader.NextResult());
            return sheetNames;
        }

        protected IExcelDataReader GetReader()
        {
            if (_reader == null)
            {
                _reader = ExcelReaderFactory.CreateReader(WorkbookStream.Reset(), new ExcelReaderConfiguration
                {
                    LeaveOpen = true,
                    Password = Password
                });
                return _reader;
            }
            _reader.Reset();
            return _reader;
        }

        protected virtual void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet name cannot be null or empty", nameof(sheetName));
        }

        protected abstract RangeReference ResolveRange(string range);
        protected abstract CellReference ResolveCell(string address);
        private int GetSheetIndex(string sheetName)
        {
            var sheetNames = GetSheetNames();
            for (int i = 0; i < sheetNames.Length; i++)
            {
                if (sheetNames[i].Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }
        public DataTable ReadRange(string sheetName, string range, bool hasHeaders, int headerRows, int rowsPerRecord)
        {
            var rangeRef = ResolveRange(range);
            var sheetIndex = GetSheetIndex(sheetName);

            var reader = GetReader();
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                FilterSheet = (_, index) => index == sheetIndex,
                UseColumnDataType = false,
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    FilterRow = (row) => rangeRef.IsRowInRange(row.Depth + 1),
                    ReadHeader = (row) =>
                    {
                        var dict = new Dictionary<int, string>();
                        
                        while (!rangeRef.IsRowInRange(row.Depth + 1))
                        {
                            row.Read();
                            headerRows++;
                            continue;
                        }

                        if (hasHeaders)
                        {
                            while (row.Depth < headerRows)
                            {
                                var index = 0;
                                for (int i = 0; i < row.FieldCount; i++)
                                {
                                    if (!rangeRef.IsColInRange(i + 1))
                                        continue;

                                    var cellValue = row.GetValue(i)?.ToString() ?? string.Empty;
                                    if (string.IsNullOrWhiteSpace(cellValue))
                                    {
                                        cellValue = $"Col{index + 1}";
                                        index++;
                                    }

                                    if (dict.ContainsKey(i))
                                        dict[i] = (dict[i] + " " + cellValue).Trim();
                                    else
                                        dict[i] = cellValue.Trim();
                                }

                                if (!row.Read())
                                    break;
                            }
                        }
                        else
                        {
                            for (int i = 0, j = 1; i < row.FieldCount; i++)
                            {
                                if (!rangeRef.IsColInRange(i + 1))
                                    continue;

                                dict.Add(i, $"Col{j}");
                                j++;
                            }
                        }

                        return dict;
                    }
                }
            });

            if (dataSet.Tables.Count > 0)
            {
                var table = dataSet.Tables[sheetName];
                var rowsToRemove = new List<DataRow>();
                if (rowsPerRecord > 1)
                {
                    DataRow recordRow = null;
                    var index = 0;
                    foreach (DataRow row in table.Rows)
                    {
                        if (index++ % rowsPerRecord == 0)
                        {
                            recordRow = row;
                            continue;
                        }

                        for (int i = 0; i < row.ItemArray.Length; i++)
                        {
                            var value = (row[i]?.ToString() ?? string.Empty).Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                recordRow[i] = $"{recordRow[i]} {value}".Trim();
                            }
                        }

                        rowsToRemove.Add(row);
                    }

                    foreach (var row in rowsToRemove)
                    {
                        table.Rows.Remove(row);
                    }
                }

                return table;
            }

            return new DataTable();
        }

        public object ReadCell(string sheetName, string address)
        {
            var cellRef = ResolveCell(address);
            var reader = GetReader();
            do
            {
                if (reader.Name == sheetName)
                {
                    var row = 0;

                    while (reader.Read())
                    {
                        ++row;
                        if (row < cellRef.Row) continue;

                        var col = cellRef.Col - 1;
                        if (col < reader.FieldCount)
                            return reader.GetValue(col);
                    }
                }
            } while (reader.NextResult());
            return null;
        }
    }
}