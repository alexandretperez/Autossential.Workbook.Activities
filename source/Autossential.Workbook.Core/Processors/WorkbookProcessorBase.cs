using Autossential.Workbook.Core.Internals;
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
        private IExcelDataReader _reader;
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
            {
                _reader?.Dispose();
                WorkbookStream?.Dispose();
            }

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

        internal virtual RangeReference ResolveRange(string range)
        {
            if (string.IsNullOrEmpty(range))
                throw new ArgumentException("Range cannot be null or empty", nameof(range));

            var rangeRef = new OpenXMLRangeReference(range);
            if (!rangeRef.IsValid)
                throw new ArgumentException("Invalid range format", nameof(range));

            return rangeRef;
        }

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

        public virtual int GetRowCount(string sheetName, string range)
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

        public virtual string[] GetSheetNames()
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

                return table;
            }

            return new DataTable();
        }

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

        private static int CountRows(IExcelDataReader reader, RangeReference rangeRef)
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
    }
}