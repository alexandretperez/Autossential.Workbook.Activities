using Autossential.Workbook.Core.Extensions;
using Autossential.Workbook.Core.Internals;
using DocumentFormat.OpenXml.Office2016.Excel;
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

            GoToSheet(reader, sheetName);

            do
            {
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
            GoToSheet(reader, sheetName);

            do
            {
                if (range == "A1" || !reader.HasRows)
                    return reader.RowCount;

                return CountRows(reader, rangeRef);

            } while (reader.NextResult());
        }

        private static void GoToSheet(ExcelDataReader reader, string sheetName)
        {
            do
            {
                if (reader.WorksheetName == sheetName)
                    break;

            } while (reader.NextResult());

            if (reader.WorksheetName != sheetName)
                throw new ArgumentException($"Sheet name '{sheetName}' was not found", nameof(sheetName));
        }

        private ExcelDataReaderOptions GetReaderOptions(string sheetName, bool hasHeaders, bool useColumnDataType)
        {
            ExcelDataReaderOptions options = null;

            if (useColumnDataType)
            {
                using var reader = GetReader();

                GoToSheet(reader, sheetName);

                var analyzer = new SchemaAnalyzer();
                var result = analyzer.Analyze(reader);

                #region Forces yes|no|y|n|t|f values to get interpreted as string instead of boolean

                var schema = result.GetSchema();
                var indexes = new List<int>();

                for (int i = 0; i < schema.Count; i++)
                {
                    var col = schema[i];
                    if (col.DataType == typeof(bool) && col.Format != null)
                        indexes.Add(i);
                }

                var builder = result.GetSchemaBuilder();
                foreach (var i in indexes)
                    builder[i].SetType(typeof(string));

                schema = builder.Build();

                #endregion

                options = new ExcelDataReaderOptions { Schema = new ExcelSchema(hasHeaders, schema.GetColumnSchema()) };

                reader.Close();

                if (options != null)
                    return options;
            }

            return new ExcelDataReaderOptions { Schema = hasHeaders ? ExcelSchema.Default : ExcelSchema.NoHeaders };
        }

        private static void AddHeadersToTable(DataTable dt, bool hasHeaders, int startCol, int endCol, int startRow, ExcelDataReader reader)
        {
            var columnSchema = reader.GetColumnSchema();

            if (hasHeaders)
            {
                if (startRow == 1)
                {
                    dt.AddColumns(startCol, endCol, (colIndex, _) => reader.GetName(colIndex), columnSchema);
                }
                else
                {
                    dt.AddColumns(startCol, endCol, (colIndex, _) => reader.GetString(colIndex), columnSchema);
                }
            }
            else
            {
                dt.AddColumns(startCol, endCol, (_, index) => $"Col{index + 1}", columnSchema);
            }
        }

        public virtual DataTable ReadRange(string sheetName, string range, bool hasHeaders, bool useColumnDataType)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);

            using var reader = GetReader(GetReaderOptions(sheetName, hasHeaders, useColumnDataType));

            GoToSheet(reader, sheetName);

            var dt = new DataTable();
            int maxColCount = 0;

            int startRow = (int)rangeRef.Start.Row;
            int endRow = (int)rangeRef.End.Row;
            int startCol = (int)rangeRef.Start.Col;
            int endCol = rangeRef.IsPartial ? reader.FieldCount : (int)rangeRef.End.Col;

            try
            {
                while (reader.Read())
                {
                    if (reader.RowNumber < startRow)
                        continue;

                    if (dt.Columns.Count == 0)
                    {
                        AddHeadersToTable(dt, hasHeaders, startCol, endCol, startRow, reader);
                        if (hasHeaders && startRow > 1)
                            if (!reader.Read())
                                break;
                    }

                    if (reader.RowNumber > endRow)
                        break;

                    maxColCount = AddRows(dt, startCol, endCol, reader, maxColCount);
                }

                if (dt.Columns.Count == 0)
                    AddHeadersToTable(dt, hasHeaders, startCol, endCol, startRow, reader);
            }
            catch (Exception)
            {
                if (!hasHeaders && useColumnDataType)
                    return ReadRange(sheetName, range, hasHeaders, false);

                throw;
            }

            return dt;
        }

        private static int AddRows(DataTable dt, int startCol, int endCol, ExcelDataReader reader, int maxColCount)
        {
            var row = dt.NewRow();

            for (int i = 0, j = startCol, k = j - 1; j <= endCol; i++, j++, k = j - 1)
            {
                var value = reader.GetValue(k);
                if (value == DBNull.Value || (value is string valStr && string.IsNullOrEmpty(valStr)))
                    continue;

                row[i] = value;
                maxColCount = Math.Max(maxColCount, i + 1);
            }

            dt.Rows.Add(row);
            return maxColCount;
        }

        //public static DataTable TrimUnusedColumns(DataTable dt, int maxColumnCount)
        //{
        //    if (maxColumnCount == 0 || dt.Columns.Count == maxColumnCount)
        //        return dt;

        //    int rows = dt.Rows.Count;

        //    object[] data = new object[rows * maxColumnCount];

        //    for (int i = 0; i < rows; i++)
        //    {
        //        var rowSpan = data.AsSpan(i * maxColumnCount, maxColumnCount);
        //        for (int j = 0; j < maxColumnCount; j++)
        //        {
        //            rowSpan[j] = dt.Rows[i][j];
        //        }
        //    }

        //    var result = new DataTable();
        //    for (int i = 0; i < maxColumnCount; i++)
        //        result.Columns.Add(dt.Columns[i].ColumnName, dt.Columns[i].DataType);

        //    for (int i = 0; i < rows; i++)
        //    {
        //        var row = result.NewRow();
        //        var rowSpan = data.AsSpan(i * maxColumnCount, maxColumnCount);
        //        for (int j = 0; j < maxColumnCount; j++)
        //        {
        //            row[j] = rowSpan[j];
        //        }
        //        result.Rows.Add(row);
        //    }

        //    return result;
        //}

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