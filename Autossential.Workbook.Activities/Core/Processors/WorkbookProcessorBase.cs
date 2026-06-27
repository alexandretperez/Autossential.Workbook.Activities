using Autossential.Workbook.Activities.Extensions;
using ExcelDataReader;
using System.Data;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal abstract class WorkbookProcessorBase : IWorkbookProcessor
    {
        public MemoryStream WorkbookStream { get; }

        public void Dispose()
        {
            if (WorkbookStream.CanRead)
                SaveInternal(WorkbookStream.ComputeHash());

            _reader?.Dispose();
            WorkbookStream?.Dispose();
        }

        public int GetColumnCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range);
            var reader = GetReader();
            int count = 0;

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                int startCol = rangeRef.Start.Col;
                int endCol = Math.Min(rangeRef.End.Col, reader.FieldCount);

                int startRow = rangeRef.Start.Row;
                int endRow = rangeRef.End.Row;

                int row = 0;
                int lastNonEmptyColumn = 0;

                while (reader.Read() && ++row <= endRow)
                {
                    if (row < startRow)
                        continue;

                    for (int i = startCol; i <= endCol; i++)
                    {
                        var value = reader.GetValue(i - 1);
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            continue;

                        lastNonEmptyColumn = i;
                    }

                    if (lastNonEmptyColumn == 0)
                        continue;

                    startCol = lastNonEmptyColumn;
                    count = lastNonEmptyColumn - (rangeRef.Start.Col - 1);
                }

                return count;
            } while (reader.NextResult());

            return count;
        }

        public int GetRowCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);
            var reader = GetReader();

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                int startCol = rangeRef.Start.Col;
                int startRow = rangeRef.Start.Row;
                int endRow = rangeRef.End.Row;
                int maxCols = Math.Min(rangeRef.End.Col, reader.FieldCount);
                int lastNonEmptyRow = 0;

                int row = 0;

                while (reader.Read() && ++row <= endRow)
                {
                    if (row < startRow)
                        continue;

                    for (int i = startCol - 1; i < maxCols; i++)
                    {
                        var value = reader.GetValue(i);
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            continue;

                        lastNonEmptyRow = row - (startRow - 1);
                        break;
                    }
                }

                return lastNonEmptyRow;
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

        public object ReadCell(string sheetName, string address)
        {
            ValidateSheetName(sheetName);
            var cellRef = ResolveCell(address);
            var reader = GetReader();
            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var row = 0;
                while (reader.Read())
                {
                    ++row;
                    if (row < cellRef.Row) continue;

                    var col = cellRef.Col - 1;
                    if (col < reader.FieldCount)
                        return reader.GetValue(col);
                }
            } while (reader.NextResult());

            return null;
        }

        public object[] ReadColumn(string sheetName, string startingCell, int limit = 0)
        {
            ValidateSheetName(sheetName);
            var reader = GetReader();

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (limit <= 0)
                    limit = int.MaxValue;

                var cell = ResolveCell(startingCell);
                var colIndex = cell.Col - 1;

                if (colIndex >= reader.FieldCount)
                    return [];

                var size = Math.Min(limit, reader.RowCount);
                var values = new object[size];
                var lastNonEmptyRow = 0;

                int index = 0;
                while (reader.Read())
                {
                    if (reader.Depth + 1 < cell.Row)
                        continue;

                    var value = reader.GetValue(colIndex);
                    values[index++] = value;

                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        lastNonEmptyRow = index;

                    if (index == limit)
                        break;
                }

                return values[..lastNonEmptyRow];
            } while (reader.NextResult());

            return [];
        }

        public DataTable ReadRange(string sheetName, string range, bool hasHeaders, int headerRows = 1, int rowsPerRecord = 1)
        {
            ValidateSheetName(sheetName);

            var reader = GetReader();
            var table = new DataTable();
            int colNameIndex = 1;

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var rangeRef = ResolveRange(range);

                var startRowIndex = rangeRef.Start.Row - 1;
                var endRowIndex = rangeRef.End.Row - 1;

                var startColIndex = rangeRef.Start.Col - 1;
                var endColIndex = rangeRef.End.Col - 1;

                endColIndex = Math.Min(endColIndex, reader.FieldCount - 1);

                int safeHeaderRows = Math.Max(headerRows, 1);
                int safeRowsPerRecord = Math.Max(rowsPerRecord, 1);

                // -- HEADERS

                if (hasHeaders)
                {
                    headerRows = safeHeaderRows;

                    var headers = new Dictionary<int, string>();
                    while (headerRows > 0 && reader.Read())
                    {
                        if (reader.Depth < startRowIndex)
                            continue;

                        headerRows--;
                        for (int i = startColIndex; i <= endColIndex; i++)
                        {
                            var name = reader.GetValue(i)?.ToString();
                            if (string.IsNullOrEmpty(name))
                                name = $"{EMPTY_COLUMN_NAME_PREFIX}{colNameIndex++}";

                            if (headers.TryGetValue(i, out string header))
                            {
                                headers[i] = $"{header} {name.Trim()}";
                                continue;
                            }

                            headers[i] = name.Trim();
                        }
                    }

                    foreach (var item in headers)
                        table.Columns.Add(item.Value, typeof(object));
                }
                else
                {
                    for (int i = startColIndex; i <= endColIndex; i++)
                    {
                        var name = $"{EMPTY_COLUMN_NAME_PREFIX}{colNameIndex++}";
                        table.Columns.Add(name, typeof(object));
                    }
                }

                // -- ROWS

                rowsPerRecord = safeRowsPerRecord;

                table.BeginLoadData();

                DataRow row = null;
                while (reader.Read() && reader.Depth <= endRowIndex)
                {
                    if (reader.Depth < startRowIndex)
                        continue;

                    row ??= table.NewRow();
                    for (int i = startColIndex, ri = 0; i <= endColIndex; i++, ri++)
                    {
                        var value = reader.GetValue(i);
                        if (row.IsNull(ri))
                        {
                            row[ri] = value is string str ? str.Trim() : value;
                            continue;
                        }

                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            continue;

                        row[ri] = $"{row[ri]} {value.ToString().Trim()}";
                    }

                    if (--rowsPerRecord <= 0)
                    {
                        table.Rows.Add(row);
                        rowsPerRecord = safeRowsPerRecord;
                        row = null;
                    }
                }

                if (row != null)
                    table.Rows.Add(row);

                table.EndLoadData();
                return table.TrimOrAppend(
                    rangeRef,
                    EMPTY_COLUMN_NAME_PREFIX,
                    colNameIndex,
                    hasHeaders,
                    safeHeaderRows,
                    safeRowsPerRecord);
            } while (reader.NextResult());

            return table;
        }

        public object[] ReadRow(string sheetName, string startingCell, int limit = 0)
        {
            ValidateSheetName(sheetName);
            var reader = GetReader();

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var cell = ResolveCell(startingCell);
                var rowIndex = cell.Row - 1;

                if (rowIndex >= reader.RowCount)
                    return [];

                limit = limit > 0 ? limit + cell.Col - 1 : int.MaxValue;

                int size = Math.Min(limit, reader.FieldCount);
                var values = new object[size];
                int lastNonEmptyColumnIndex = -1;
                int colIndex = cell.Col - 1;

                while (reader.Read())
                {
                    if (reader.Depth < rowIndex)
                        continue;

                    for (int ci = colIndex, vi = 0; ci < size; ci++, vi++)
                    {
                        var value = reader.GetValue(ci);
                        values[vi] = value;
                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                            lastNonEmptyColumnIndex = ci;
                    }

                    lastNonEmptyColumnIndex++;
                    break;
                }

                size = lastNonEmptyColumnIndex - colIndex;
                return values[..size];
            } while (reader.NextResult());

            return [];
        }

        public void Save()
        {
            var computedHash = WorkbookStream.ComputeHash();
            if (_lastSaveHash == computedHash)
                return;

            SaveInternal(_lastSaveHash);
            _lastSaveHash = computedHash;
            WorkbookHash = computedHash;
        }

        public abstract void WriteCell(string sheetName, string address, object value);

        public abstract void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders);

        protected WorkbookProcessorBase(string filePath, string password)
        {
            FilePath = filePath;
            Password = password;
            WorkbookStream = new MemoryStream();

            if (File.Exists(FilePath))
            {
                var bytes = File.ReadAllBytes(FilePath);
                WorkbookStream.Write(bytes, 0, bytes.Length);
                WorkbookStream.Position = 0;
            }
            else
            {
                CreateNew();
                WorkbookStream.Position = 0;
            }

            WorkbookHash = WorkbookStream.ComputeHash();
        }

        protected string FilePath { get; }
        protected string Password { get; }
        protected string WorkbookHash { get; set; }

        protected abstract void CreateNew();

        protected IExcelDataReader GetReader()
        {
            if (_reader == null)
            {
                WorkbookStream.Position = 0;
                _reader = ExcelReaderFactory.CreateReader(WorkbookStream, new ExcelReaderConfiguration
                {
                    LeaveOpen = true,
                    Password = Password
                });
                return _reader;
            }
            _reader.Reset();
            return _reader;
        }

        protected abstract CellReference ResolveCell(string address);

        protected abstract RangeReference ResolveRange(string range);

        protected virtual void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet name cannot be null or empty", nameof(sheetName));
        }

        private const string EMPTY_COLUMN_NAME_PREFIX = "Col";
        private string _lastSaveHash = null;
        private IExcelDataReader _reader;

        private void SaveInternal(string computedHash)
        {
            if (computedHash != WorkbookHash)
            {
                WorkbookStream.Position = 0;
                using var fs = File.Create(FilePath);
                WorkbookStream.CopyTo(fs, WorkbookStream.CalculateBufferSize());
                WorkbookHash = computedHash;
            }
        }

        private string _defaultSheetName;

        protected string SetDefaultSheetName(string sheetName = "Sheet1") => (_defaultSheetName = sheetName);

        protected string ConsumeDefaultSheetName()
        {
            string name = _defaultSheetName;
            _defaultSheetName = null;
            return name;
        }

        public (string, int, int) FindValue(string sheetName, string range, object value)
        {
            ValidateSheetName(sheetName);
            var reader = GetReader();

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var rangeRef = ResolveRange(range);

                var startColIndex = rangeRef.Start.Col - 1;
                var endCol = Math.Min(rangeRef.End.Col, reader.FieldCount);

                while (reader.Read())
                {
                    if (reader.Depth + 1 < rangeRef.Start.Row)
                        continue;

                    for (int i = startColIndex; i < endCol; i++)
                    {
                        var cellValue = reader.GetValue(i);
                        if (cellValue is null || string.IsNullOrEmpty(cellValue.ToString()))
                        {
                            if (value is null || string.IsNullOrEmpty(value.ToString()))
                            {
                                cellValue = null;
                                value = null;
                            }
                        }

                        if (cellValue == value || cellValue?.ToString() == value?.ToString())
                        {
                            var cell = new CellReference(i + 1, reader.Depth + 1);
                            return (cell.ToString(), cell.Col, cell.Row);
                        }
                    }
                }

                break;
            } while (reader.NextResult());

            return (string.Empty, -1, -1);
        }

        public abstract void DeleteSheet(string sheetName);

        public abstract void InsertSheet(string sheetName, int? position);

        public abstract void RenameSheet(string fromSheetName, string toSheetName);
    }
}