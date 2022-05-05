using Autossential.Workbook.Core.Enums;
using Autossential.Workbook.Core.Internals;
using ExcelDataReader;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;

namespace Autossential.Workbook.Core.Adapters
{
    public abstract class WorkbookAdapterBase : IWorkbookAdapter
    {
        private IExcelDataReader _reader;
        private bool _requiresSave;

        public WorkbookAdapterBase(string filePath)
        {
            FilePath = filePath;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            OpenFile();
        }

        protected string FilePath { get; private set; }

        protected bool IsNewWorkbook { get; private set; }

        protected Stream WorkbookFileStream { get; private set; }

        public abstract void AddHyperLink(string sheetName, string cell, string label, string link, string tooltip);

        public virtual void Save()
        {
            if (_requiresSave)
                GetSaveHandler().Invoke();
        }
        public abstract void CreateNew();

        public void Dispose()
        {
            Dispose(true);

            _reader?.Dispose();

            if (WorkbookFileStream != null)
            {
                WorkbookFileStream.Close();
                WorkbookFileStream.Dispose();
            }
        }

        public abstract Action GetSaveHandler();

        public abstract void Dispose(bool disposing);

        public abstract void FreezePanes(string sheetName, int cols, int rows);

        public abstract string[] GetHyperlinks(string sheetName, string range);

        public string[] GetSheetNames()
        {
            var reader = GetExcelReader();
            var sheetNames = new string[reader.ResultsCount];
            var i = 0;
            do
            {
                sheetNames[i++] = reader.Name;
            } while (reader.NextResult());
            return sheetNames;
        }

        public DataTable ReadRange(string sheetName, string range, bool addHeaders)
        {
            var reader = GetExcelReader();

            var dt = new DataTable();
            var addr = new RangeAddress(range);
            if (!addr.First.IsValid)
                throw new ArgumentException("The range is not valid " + range, nameof(range));

            do
            {
                if (reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    addr.Last.SetDefault(reader.FieldCount, reader.RowCount);

                    var headers = GetHeaderValues(reader, addHeaders, addr);
                    var values = GetDataValues(reader, addr, out Type[] colTypes);
                    TrimEnd(ref values);
                    BuildDataTable(dt, headers, values, colTypes);
                    break;
                }
            } while (reader.NextResult());

            return dt;
        }

        public abstract int RemoveHyperLinks(string sheetName, string range);

        public void RequiresSave()
        {
            _requiresSave = true;
        }

        public abstract void WriteCell(string sheetName, string cell, object value);

        public abstract void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders);

        private static void BuildDataTable(DataTable dt, string[] headers, object[][] values, Type[] colTypes)
        {
            int colIndex = 0;
            for (int i = 0; i < headers.Length; i++)
                dt.Columns.Add(GetColumnName(headers[i], ref colIndex), colTypes[i] ?? typeof(object));

            dt.BeginLoadData();

            foreach (var value in values)
                dt.LoadDataRow(value, LoadOption.OverwriteChanges);

            dt.EndLoadData();
        }

        private static string GetColumnName(string value, ref int index)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                index++;
                return "Column" + index;
            }

            return value;
        }

        private static object[][] GetDataValues(IExcelDataReader reader, RangeAddress ra, out Type[] types)
        {
            types = new Type[ra.ColsUsed()];
            var values = new object[ra.RowsUsed()][];
            int i = 0;
            while (reader.Read())
            {
                values[i] = new object[ra.ColsUsed()];
                for (int j = ra.First.Col - 1, k = 0; j < ra.Last.Col; j++, k++)
                {
                    var v = reader.GetValue(j);
                    values[i][k] = v;

                    var colType = types[k];
                    if (v == null || colType == typeof(string) || colType == typeof(object))
                        continue;

                    var type = v.GetType();

                    if (type == typeof(double))
                    {
                        type = colType ?? typeof(int);

                        if (colType == null || colType == typeof(int))
                        {
                            if (((double)v % 1) > double.Epsilon)
                                type = typeof(decimal);
                        }
                    }
                    else if (type != colType && colType != null)
                    {
                        type = typeof(object);
                    }

                    types[k] = type;
                }
                i++;
            }

            return values;
        }

        private static string[] GetHeaderValues(IExcelDataReader reader, bool addHeaders, RangeAddress ra)
        {
            var headers = new string[ra.ColsUsed()];
            if (addHeaders)
            {
                var isValid = false;
                var firstRow = ra.First.Row;
                var firstCol = ra.Last.Col; // starts with the greater one
                var rowIndex = 1;
                while (reader.Read())
                {
                    rowIndex++;

                    if (--firstRow > 0) continue;

                    for (int i = ra.First.Col - 1, j = 0; i < ra.Last.Col; i++, j++)
                    {
                        var v = reader.GetValue(i);
                        if (v == null)
                            continue;

                        firstCol = Math.Min(firstCol, j + 1);
                        headers[j] = v.ToString();
                        isValid = true;
                    }

                    if (isValid) break;
                }
                ra.First.Override(firstCol, rowIndex);

                if (firstCol > 0)
                {
                    var result = new string[headers.Length - firstCol + 1];
                    Array.Copy(headers, firstCol - 1, result, 0, result.Length);
                    return result;
                }
            }

            return headers;
        }

        private static void TrimEnd(ref object[][] values)
        {
            var emptyCounter = 0;
            foreach (var value in values)
            {
                var empty = true;
                foreach (var v in value)
                {
                    if (v != null)
                    {
                        empty = false;
                        break;
                    }
                }
                if (empty)
                {
                    emptyCounter++;
                    continue;
                }
                emptyCounter = 0;
            }

            if (emptyCounter > 0)
                Array.Resize(ref values, values.Length - emptyCounter);
        }

        private IExcelDataReader GetExcelReader()
        {
            if (_reader == null)
            {
                _reader = ExcelReaderFactory.CreateReader(WorkbookFileStream, new ExcelReaderConfiguration
                {
                    LeaveOpen = true
                });
            }
            else
            {
                _reader.Reset();
            }

            WorkbookFileStream.Position = 0;
            return _reader;
        }

        private void OpenFile()
        {
            if (!File.Exists(FilePath))
            {
                CreateNew();
                IsNewWorkbook = true;
            }

            WorkbookFileStream = new FileStream(FilePath, FileMode.Open);
        }

        public abstract void AppendRange(string sheetName, DataTable dataTable);
        public abstract void DrawBorder(string sheetName, string range, Border border, BorderStyle style, Color color);
        public abstract void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation);
    }
}