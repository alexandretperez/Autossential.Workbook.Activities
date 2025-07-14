using Autossential.Workbook.Core.Extensions;
using Autossential.Workbook.Core.Internals;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
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

            if (File.Exists(filePath))
            {
                var bytes = File.ReadAllBytes(FilePath);
                WorkbookStream.Write(bytes, 0, bytes.Length);
                WorkbookStream.Reset();
            }
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
                if (reader.Name == sheetName)
                {
                    CountColumns(reader, rangeRef, processedColumns);
                    break;
                }

            } while (reader.NextResult());

            return processedColumns.Count;
        }

        public string[] GetSheetNames()
        {
            var reader = GetReader();
            var sheetNames = new string[reader.ResultsCount];
            int i = 0;
            do
            {
                sheetNames[i++] = reader.Name;
            }
            while (reader.NextResult());
            return sheetNames;
        }

        public int GetRowCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);
            int count = 0;
            var reader = GetReader();

            do
            {
                if (reader.Name == sheetName)
                {
                    count = CountRows(reader, rangeRef);
                    break;
                }
            }
            while (reader.NextResult());
            return count;
        }


        public virtual DataTable ReadRange(string sheetName, string range, bool hasHeaders, bool useColumnDataType)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);
            var index = Array.IndexOf(GetSheetNames(), sheetName);
            if (index == -1)
                throw new ArgumentException("Sheet name not found", nameof(sheetName));

            var reader = GetReader();

            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                FilterSheet = (_, sheetIndex) => sheetIndex == index,
                UseColumnDataType = useColumnDataType,
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = hasHeaders,
                    FilterRow = (rowReader) => rangeRef.IsRowInRange((uint)rowReader.Depth + 1),
                    FilterColumn = (_, columnIndex) => rangeRef.IsColInRange((uint)columnIndex + 1)
                }
            });

            if (dataSet.Tables.Count > 0)
            {
                var table = dataSet.Tables[sheetName];
                if (!hasHeaders)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        table.Columns[i].ColumnName = "Col" + (i + 1);
                    }
                }

                if (hasHeaders && rangeRef.Start.Row > 1)
                {
                    var firstRow = table.Rows[0];
                    var colNum = 1;
                    foreach (DataColumn col in table.Columns)
                    {
                        var colName = firstRow[col.Ordinal].ToString();
                        if (string.IsNullOrEmpty(colName))
                        {
                            do
                            {
                                colName = "Col" + (colNum++);

                            } while (firstRow.ItemArray.Contains(colName));
                        }

                        col.ColumnName = colName;
                    }
                    table.Rows.Remove(firstRow);
                }

                return table;
            }

            return new DataTable();
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

        protected IExcelDataReader GetReader() =>
            ExcelReaderFactory.CreateReader(WorkbookStream.Reset(), new ExcelReaderConfiguration { LeaveOpen = true });

        private static void CountColumns(IExcelDataReader reader, RangeReference rangeRef, HashSet<int> processedColumns)
        {
            var row = 0;
            var emptyColumns = new HashSet<int>();

            while (reader.Read())
            {
                ++row;
                if (row < rangeRef.Start.Row) continue;
                if (row > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, reader.FieldCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    if (processedColumns.Contains(col))
                        continue;

                    if (reader.GetValue(col - 1) != null)
                    {
                        processedColumns.Add(col);

                        foreach (var c in emptyColumns.Where(emptyCol => emptyCol < col))
                            processedColumns.Add(c);

                        emptyColumns.RemoveWhere(c => c <= col);

                        if (processedColumns.Count == (size - (rangeRef.Start.Col - 1)))
                            return;

                        continue;
                    }

                    emptyColumns.Add(col);
                }
            }
        }

        internal static int CountRows(IExcelDataReader reader, RangeReference rangeRef)
        {
            int row = 0;
            var emptyRowCount = 0;
            int count = 0;

            while (reader.Read())
            {
                ++row;
                if (row < rangeRef.Start.Row) continue;
                if (row > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, reader.FieldCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    if (reader.GetValue(col - 1) != null)
                    {
                        count += emptyRowCount + 1;
                        emptyRowCount = 0;
                        size = -1;

                        break;
                    }
                }

                if (size != -1)
                    emptyRowCount++;
            }

            return count;
        }

        public abstract void RenameSheet(int sheetIndex, string newSheetName);
        public abstract void RenameSheet(string fromSheetName, string toSheetName);
        public abstract void DeleteSheet(string sheetName);
        public abstract void ActivateSheet(string sheetName);
        public abstract void ActivateSheet(int sheetIndex);
        public abstract (int index, string name) GetActiveSheet();
        public abstract void WriteRange(DataTable dt, string sheetName, string startCell, bool addHeaders);
    }
}