using Autossential.Workbook.Core.Internals;
using Sylvan.Data;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Processors
{
    public abstract class WorkbookProcessorBase : IWorkbookProcessor
    {
        protected WorkbookProcessorBase(string filePath)
        {
            FilePath = filePath;

            WorkbookStream = new MemoryStream();

            var bytes = File.ReadAllBytes(FilePath);
            WorkbookStream.Write(bytes, 0, bytes.Length);
            WorkbookStream.Reset();
        }

        protected MemoryStream WorkbookStream { get; }

        public string FilePath { get; }

        public abstract void Save();

        protected bool RequiresSave { get; set; }

        private bool _disposed;

        public virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
                WorkbookStream?.Dispose();

            _disposed = true;
        }

        public void Dispose()
        {
            try
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            catch (Exception e)
            {
                Trace.WriteLine(e.Message);
            }
        }

        ~WorkbookProcessorBase()
        {
            Dispose(false);
        }

        protected virtual void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet name cannot be null or empty", nameof(sheetName));
        }

        internal abstract RangeReference ResolveRange(string range);

        public virtual int GetColumnCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);
            var processedColumns = new HashSet<int>();
            var reader = GetReader();

            do
            {
                if (reader.WorksheetName != sheetName)
                    continue;

                CountColumns(reader, rangeRef, processedColumns);
                break;

            } while (reader.NextResult());

            return processedColumns.Count;
        }

        public string[] GetSheetNames()
        {
            using var reader = GetReader();
            return reader.WorksheetNames.ToArray();
        }

        public int GetRowCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range);

            using var reader = GetReader();
            do
            {
                if (reader.WorksheetName != sheetName)
                    continue;

                if (range == "A1" || !reader.HasRows)
                    return reader.RowCount;

                return CountRows(reader, rangeRef);

            } while (reader.NextResult());

            return 0;
        }

        private ExcelDataReaderOptions GetReaderOptions(string sheetName, bool hasHeaders, bool useColumnDataType)
        {
            ExcelDataReaderOptions options = null;

            if (useColumnDataType)
            {
                using var reader = GetReader();

                do
                {
                    if (reader.WorksheetName == sheetName)
                    {
                        var analyzer = new SchemaAnalyzer();
                        var result = analyzer.Analyze(reader);
                        options = new ExcelDataReaderOptions { Schema = new ExcelSchema(hasHeaders, result.GetSchema().GetColumnSchema()) };
                    }

                } while (reader.NextResult());

                reader.Close();

                if (options != null)
                    return options;
            }

            return new ExcelDataReaderOptions { Schema = hasHeaders ? ExcelSchema.Default : ExcelSchema.NoHeaders };
        }

        public virtual DataTable ReadRange(string sheetName, string range, bool hasHeaders, bool useColumnDataType)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);

            using var reader = GetReader(GetReaderOptions(sheetName, hasHeaders, useColumnDataType));

            do
            {
                if (reader.WorksheetName == sheetName)
                    break;

            } while (reader.NextResult());

            if (reader.WorksheetName != sheetName)
                throw new ArgumentException("Sheet name not found", nameof(sheetName));

            var dt = new DataTable();

            if (range == "A1")
                return ReadByLoadReader(sheetName, range, hasHeaders, useColumnDataType, reader, dt);


            int maxColumnCount = 1;

            while (reader.Read())
            {
                if (!rangeRef.IsRowInRange(reader.RowNumber))
                {
                    if (reader.RowNumber > rangeRef.End.Row)
                        break;

                    continue;
                }

                if (dt.Columns.Count == 0)
                {
                    AddHeaders(hasHeaders, rangeRef, reader, dt);
                    if (rangeRef.Start.Row > 1 && hasHeaders)
                        continue;
                }

                if (reader.RowNumber == 1)
                    AddRows(dt, reader, rangeRef, (i) => reader.GetName(i), ref maxColumnCount);
                else
                    AddRows(dt, reader, rangeRef, (i) => reader.GetValue(i), ref maxColumnCount);
            }

            return rangeRef.IsPartial ? TrimUnusedColumns(dt, maxColumnCount) : dt;
        }

        public static DataTable TrimUnusedColumns(DataTable dt, int maxColumnCount)
        {
            if (dt.Columns.Count == maxColumnCount)
                return dt;

            int rows = dt.Rows.Count;

            object[] data = new object[rows * maxColumnCount];

            for (int i = 0; i < rows; i++)
            {
                var rowSpan = data.AsSpan(i * maxColumnCount, maxColumnCount);
                for (int j = 0; j < maxColumnCount; j++)
                {
                    rowSpan[j] = dt.Rows[i][j];
                }
            }

            var result = new DataTable();
            for (int i = 0; i < maxColumnCount; i++)
                result.Columns.Add(dt.Columns[i].ColumnName, dt.Columns[i].DataType);

            for (int i = 0; i < rows; i++)
            {
                var row = result.NewRow();
                var rowSpan = data.AsSpan(i * maxColumnCount, maxColumnCount);
                for (int j = 0; j < maxColumnCount; j++)
                {
                    row[j] = rowSpan[j];
                }
                result.Rows.Add(row);
            }

            return result;
        }

        private static void AddRows(DataTable dt, ExcelDataReader reader, RangeReference rangeRef, Func<int, object> readValue, ref int maxColumnCount)
        {
            var maxColIndex = 0;
            var row = dt.NewRow();
            rangeRef.ForEachColumn((col, index) =>
            {
                var value = readValue(col - 1);
                if (value == DBNull.Value || (value is string valueStr && string.IsNullOrEmpty(valueStr)))
                    return;

                row[index] = value;
                maxColIndex = Math.Max(maxColIndex, index);
            });
            dt.Rows.Add(row);
            maxColumnCount = Math.Max(maxColumnCount, maxColIndex + 1);
        }

        private static void AddHeaders(bool hasHeaders, RangeReference rangeRef, ExcelDataReader reader, DataTable dt)
        {
            if (hasHeaders)
            {
                if (rangeRef.Start.Row == 1)
                {
                    rangeRef.ForEachColumn((col, _) => dt.Columns.Add(reader.GetName(col - 1)));
                }
                else
                {
                    rangeRef.ForEachColumn((col, _) => dt.Columns.Add(reader.GetString(col - 1)));
                }

                return;
            }

            rangeRef.ForEachColumn((_, index) => dt.Columns.Add($"Col{index + 1}"));
        }

        private DataTable ReadByLoadReader(string sheetName, string range, bool hasHeaders, bool useColumnDataType, ExcelDataReader reader, DataTable dt)
        {
            try
            {
                dt.Load(reader);
                if (!hasHeaders)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                        dt.Columns[i].ColumnName = $"Col{i + 1}";
                }
            }
            catch (Exception)
            {
                reader.Close();

                // retry disabling the data type constraints

                if (useColumnDataType)
                    return ReadRange(sheetName, range, hasHeaders, false);
            }

            return dt;
        }

        protected abstract ExcelDataReader GetReader(ExcelDataReaderOptions options = null);

        private static void CountColumns(ExcelDataReader reader, RangeReference rangeRef, HashSet<int> processedColumns)
        {
            var emptyColumns = new HashSet<int>();

            int colsCount = Math.Max(0, reader.FieldCount);
            if (colsCount == 0)
                colsCount = (int)rangeRef.End.Col;

            do
            {
                if (reader.RowNumber < rangeRef.Start.Row) continue;
                if (reader.RowNumber > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, colsCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    if (processedColumns.Contains(col))
                        continue;

                    var value = reader.RowNumber == 1 ? reader.GetName(col - 1) : reader.GetValue(col - 1);

                    if (value == null || value == DBNull.Value || (value is string valueStr && string.IsNullOrEmpty(valueStr)))
                    {
                        emptyColumns.Add(col);
                        continue;
                    }

                    processedColumns.Add(col);

                    foreach (var c in emptyColumns.Where(emptyCol => emptyCol < col))
                        processedColumns.Add(c);

                    emptyColumns.RemoveWhere(c => c <= col);

                    if (processedColumns.Count == (size - (rangeRef.Start.Col - 1)))
                        return;
                }
            } while (reader.Read());
        }

        internal static int CountRows(ExcelDataReader reader, RangeReference rangeRef)
        {
            var emptyRowCount = 0;
            int count = 0;

            int colsCount = Math.Max(0, reader.FieldCount);
            if (colsCount == 0)
                colsCount = (int)rangeRef.End.Col;

            do
            {
                if (reader.RowNumber < rangeRef.Start.Row) continue;
                if (reader.RowNumber > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, colsCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    var value = reader.RowNumber == 1 ? reader.GetName(col - 1) : reader.GetValue(col - 1);

                    if (value == null || value == DBNull.Value || (value is string valueStr && string.IsNullOrEmpty(valueStr)))
                        continue;

                    count += emptyRowCount + 1;
                    emptyRowCount = 0;
                    size = -1;

                    break;
                }

                if (size != -1)
                    emptyRowCount++;

            } while (reader.Read());

            return count;
        }

        public abstract void RenameSheet(int sheetIndex, string newSheetName);
        public abstract void RenameSheet(string fromSheetName, string toSheetName);
        public abstract void DeleteSheet(string sheetName);
        public abstract void ActivateSheet(string sheetName);
        public abstract void ActivateSheet(int sheetIndex);
        public abstract (int index, string name) GetActiveSheet();
    }
}