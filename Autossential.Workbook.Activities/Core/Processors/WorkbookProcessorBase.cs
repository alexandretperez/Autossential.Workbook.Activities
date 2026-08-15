using Autossential.Workbook.Activities.Extensions;
using ExcelDataReader;
using System.Data;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal abstract class WorkbookProcessorBase : IWorkbookProcessor
    {
        public abstract bool IsBIFF8 { get; }
        public abstract bool IsOpenXML { get; }
        public MemoryStream WorkbookStream { get; }

        public abstract void DeleteSheet(string sheetName);

        public void Dispose()
        {
            if (WorkbookStream.CanRead)
                SaveInternal(WorkbookStream.ComputeHash());

            _reader?.Dispose();
            WorkbookStream?.Dispose();
        }

        public (string, int, int) FindValue(string sheetName, string range, object value)
        {
            ValidateSheetName(sheetName);
            var reader = GetReader();

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var rangeRef = ResolveRange(range).Normalize(IsOpenXML ? CellRef.MaxOpenXML() : CellRef.MaxBIFF8());

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
                            int col = i + 1;
                            int row = reader.Depth + 1;
                            var address = $"{CellRef.GetColumnName(col)}{row}";
                            return (address, col, row);
                        }
                    }
                }

                break;
            } while (reader.NextResult());

            return (string.Empty, -1, -1);
        }

        public abstract void FreezePanes(string sheetName, int colsToFreeze, int rowsToFreeze);

        public int GetColumnCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range).Normalize(IsOpenXML ? CellRef.MaxOpenXML() : CellRef.MaxBIFF8());
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

            var rangeRef = ResolveRange(range).Normalize(IsOpenXML ? CellRef.MaxOpenXML() : CellRef.MaxBIFF8());
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

        public abstract void HideSheet(string sheetName);

        public abstract void InsertSheet(string sheetName, int? position);

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
                var normRangeRef = rangeRef.Normalize(IsOpenXML ? CellRef.MaxOpenXML() : CellRef.MaxBIFF8());

                var startRowIndex = normRangeRef.Start.Row - 1;
                var endRowIndex = normRangeRef.End.Row - 1;

                var startColIndex = normRangeRef.Start.Col - 1;
                var endColIndex = normRangeRef.End.Col - 1;

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
                                name = null;

                            if (headers.TryGetValue(i, out string header))
                            {
                                if (name is null)
                                    continue;

                                headers[i] = $"{header} {name.Trim()}";
                                continue;
                            }

                            var autoName = $"{EMPTY_COLUMN_NAME_PREFIX}{colNameIndex}";
                            headers[i] = (name ?? autoName).Trim();

                            if (string.Equals(headers[i], autoName, StringComparison.OrdinalIgnoreCase))
                                colNameIndex++;
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

                        var currValue = row[ri]?.ToString();
                        if (string.IsNullOrEmpty(currValue))
                        {
                            row[ri] = value.ToString().Trim();
                            continue;
                        }

                        row[ri] = $"{currValue} {value.ToString().Trim()}";
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

        public abstract void RenameSheet(string fromSheetName, string toSheetName);

        public void Save()
        {
            var computedHash = WorkbookStream.ComputeHash();
            if (_lastSaveHash == computedHash)
                return;

            SaveInternal(computedHash);
            _lastSaveHash = computedHash;
            WorkbookHash = computedHash;
        }

        public abstract void UnhideSheet(string sheetName);

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

        protected string ConsumeDefaultSheetName()
        {
            string name = _defaultSheetName;
            _defaultSheetName = null;
            return name;
        }

        protected abstract void CreateNew();

        protected IExcelDataReader GetReader()
        {
            var currentHash = WorkbookStream.ComputeHash();

            if (_reader == null || currentHash != _readerStreamHash)
            {
                _reader?.Dispose();
                WorkbookStream.Position = 0;
                _reader = ExcelReaderFactory.CreateReader(WorkbookStream, new ExcelReaderConfiguration
                {
                    LeaveOpen = true,
                    Password = Password
                });
                _readerStreamHash = currentHash;
            }
            else
            {
                _reader.Reset();
            }

            return _reader;
        }

        protected abstract CellRef ResolveCell(string address);

        protected abstract RangeRef ResolveRange(string address);

        protected string SetDefaultSheetName(string sheetName = "Sheet1") => (_defaultSheetName = sheetName);

        protected virtual void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet name cannot be null or empty", nameof(sheetName));
        }

        private const string EMPTY_COLUMN_NAME_PREFIX = "Col";

        private string _defaultSheetName;

        private string _lastSaveHash = null;

        private IExcelDataReader _reader;

        private string _readerStreamHash;

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

        public void AppendRange(string sheetName, DataTable data)
        {
            ValidateSheetName(sheetName);

            var reader = GetReader();
            int col = 0;
            int row = 0;

            do
            {
                if (!reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var endCol = reader.FieldCount;
                while (reader.Read())
                {
                    for (int i = 0; i < endCol; i++)
                    {
                        var value = reader.GetValue(i);
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            continue;

                        row = reader.Depth + 1;

                        if (col > 0)
                            break;

                        col = i + 1;
                        break;
                    }
                }
            } while (reader.NextResult());

            col = Math.Max(col, 1);
            row = Math.Max(row + 1, 1);

            WriteRange(sheetName, data, new CellRef(col, row).GetAddress(), false);
        }

        private static HashSet<int> ResolveRowsOrColumnsReferences(ReadOnlySpan<char> span, bool deletingRows)
        {
            var positions = new HashSet<int>();

            int i = 0;
            int len = span.Length;

            int start = -1;
            int end = -1;
            int init = -1;

            while (i < len)
            {
                var c = span[i];
                if (c == ' ')
                {
                    i++;
                    continue;
                }

                if (start == -1)
                    start = i;

                if (c == ',' || c == ':')
                    end = i;

                if (i + 1 == len)
                    end = len;

                if (end > 0)
                {
                    var slice = span[start..end];
                    int value;
                    if (deletingRows)
                    {
                        if (!int.TryParse(slice, out value) || value < 1)
                            throw new FormatException($"Invalid row number '{slice}' in '{span}'.");
                    }
                    else
                    {
                        value = CellRef.GetColumnIndex(slice.TrimEnd());
                    }

                    if (init > -1)
                    {
                        if (c == ':')
                            throw new FormatException($"Invalid double range in '{span}'.");

                        if (value < init)
                        {
                            if (deletingRows)
                                throw new FormatException($"Invalid '{init}:{value}' range: end is less than start.");
                            else
                                throw new FormatException($"Invalid '{CellRef.GetColumnName(init)}:{CellRef.GetColumnName(value)}' range: end is less than start.");
                        }

                        while (++init <= value)
                        {
                            positions.Add(init);
                        }
                        init = -1;
                    }
                    else
                    {
                        if (c == ':')
                            init = value;

                        positions.Add(value);
                    }

                    start = -1;
                    end = -1;
                }

                i++;
            }

            return positions;
        }
        protected static HashSet<int> ResolveRowsReferences(string references)
        {
            if (string.IsNullOrWhiteSpace(references))
                throw new ArgumentException("Row references cannot be empty.");

            return ResolveRowsOrColumnsReferences(references, true);
        }

        protected static HashSet<int> ResolveColumnsReferences(string references)
        {
            if (string.IsNullOrWhiteSpace(references))
                throw new ArgumentException("Column references cannot be empty.");

            return ResolveRowsOrColumnsReferences(references, false);
        }

        public abstract void DeleteColumns(string sheetName, string references);
        public abstract void DeleteRows(string sheetName, string references);
    }
}