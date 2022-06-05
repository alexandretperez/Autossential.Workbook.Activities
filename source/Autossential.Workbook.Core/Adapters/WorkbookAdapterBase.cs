using Autossential.Workbook.Core.Enums;
using Autossential.Workbook.Core.Internals;
using ExcelDataReader;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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

            OpenOrCreate();
        }

        protected string FilePath { get; private set; }

        protected bool IsNewWorkbook { get; private set; }

        protected Stream WorkbookFileStream { get; private set; }
        public abstract int MaxRows { get; }
        public abstract bool IsOpenXml { get; }

        public void AddHyperLink(string sheetName, string cell, string label, string link, string tooltip)
        {
            var sheetCell = GetOrCreateCell(sheetName, cell);
            var linkType = HyperlinkType.Document;

            if (Regex.IsMatch(link, "(https?|ftp)://", RegexOptions.IgnoreCase))
            {
                linkType = HyperlinkType.Url;
            }
            else if (link.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            {
                linkType = HyperlinkType.Email;
            }
            else if (!string.IsNullOrEmpty(Path.GetExtension(link)))
            {
                linkType = HyperlinkType.File;
            }

            sheetCell.Hyperlink = GetWorkbook().GetCreationHelper().CreateHyperlink(linkType);
            sheetCell.Hyperlink.Address = link;
            if (string.IsNullOrEmpty(label))
                label = link;

            sheetCell.SetCellValue(label);
            RequiresSave();
        }

        public void AppendRange(string sheetName, DataTable dataTable)
        {
            var sheet = GetOrCreateSheet(sheetName);
            var cell = "A1";
            var row = sheet.GetRow(sheet.LastRowNum);
            if (row != null)
                cell = new CellReference(row.RowNum + 1, row.FirstCellNum).FormatAsString();

            WriteRange(sheetName, cell, dataTable, false);
        }

        public abstract void CreateNew();

        public abstract void Dispose(bool disposing);

        public void Dispose()
        {
            try
            {
                Dispose(true);
                _reader?.Dispose();

                if (WorkbookFileStream != null)
                    WorkbookFileStream.Dispose();
            }
            catch (Exception e)
            {
                Trace.WriteLine(e.Message);
            }
        }

        protected NPOI.SS.UserModel.BorderStyle ConvertBorderStyle(Enums.BorderStyle style)
        {
            return (NPOI.SS.UserModel.BorderStyle)Enum.Parse(typeof(NPOI.SS.UserModel.BorderStyle), style.ToString());
        }

        public void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
        {
            var wb = GetWorkbook();
            var sheet = GetOrCreateSheet(sheetName);
            var borderStyle = ConvertBorderStyle(style);
            var borderColor = ConvertColor(color);

            if (border == Border.Outside)
            {
                var addr = new Internals.RangeAddress(range, true);
                var row = sheet.Row(addr.First.Row);
                row.Cell(addr.First.Col).CellStyle = CreateBorderStyle(wb, "TL", borderStyle, borderColor);
                row.Cell(addr.Last.Col).CellStyle = CreateBorderStyle(wb, "TR", borderStyle, borderColor);

                row = sheet.Row(addr.Last.Row);
                row.Cell(addr.First.Col).CellStyle = CreateBorderStyle(wb, "BL", borderStyle, borderColor);
                row.Cell(addr.Last.Col).CellStyle = CreateBorderStyle(wb, "BR", borderStyle, borderColor);

                var left = CreateBorderStyle(wb, "L", borderStyle, borderColor);
                var right = CreateBorderStyle(wb, "R", borderStyle, borderColor);
                for (int i = addr.First.Row + 1; i < addr.Last.Row; i++)
                {
                    row = sheet.Row(i);
                    row.Cell(addr.First.Col).CellStyle = left;
                    row.Cell(addr.Last.Col).CellStyle = right;
                }

                var top = CreateBorderStyle(wb, "T", borderStyle, borderColor);
                var bottom = CreateBorderStyle(wb, "B", borderStyle, borderColor);

                var rowTop = sheet.Row(addr.First.Row);
                var rowBottom = sheet.Row(addr.Last.Row);
                for (int i = addr.First.Col + 1; i < addr.Last.Col; i++)
                {
                    rowTop.Cell(i).CellStyle = top;
                    rowBottom.Cell(i).CellStyle = bottom;
                }
            }
            else if (border == Border.Inside)
            {
                var addr = new Internals.RangeAddress(range, true);
                var br = CreateBorderStyle(wb, "BR", borderStyle, borderColor);
                var b = CreateBorderStyle(wb, "B", borderStyle, borderColor);
                var r = CreateBorderStyle(wb, "R", borderStyle, borderColor);

                foreach (var cell in GetOrCreateCells(sheetName, range))
                {
                    if (cell.RowIndex == addr.Last.Row && cell.ColumnIndex == addr.Last.Col)
                        break;

                    cell.CellStyle = cell.RowIndex == addr.Last.Row
                        ? r
                        : cell.ColumnIndex == addr.Last.Col
                            ? b
                            : br;
                }
            }
            else
            {
                ICellStyle cellStyle = null;
                switch (border)
                {
                    case Border.None:
                    case Border.All:
                        cellStyle = CreateBorderStyle(wb, "TBLR", borderStyle, borderColor);
                        break;
                    case Border.Top:
                        cellStyle = CreateBorderStyle(wb, "T", borderStyle, borderColor);
                        break;
                    case Border.Bottom:
                        cellStyle = CreateBorderStyle(wb, "B", borderStyle, borderColor);
                        break;
                    case Border.Left:
                        cellStyle = CreateBorderStyle(wb, "L", borderStyle, borderColor);
                        break;
                    case Border.Right:
                        cellStyle = CreateBorderStyle(wb, "R", borderStyle, borderColor);
                        break;
                }

                foreach (var cell in GetOrCreateCells(sheetName, range))
                    cell.CellStyle = cellStyle;
            }

            RequiresSave();
        }

        protected abstract ICellStyle CreateBorderStyle(IWorkbook workbook, string anchors, NPOI.SS.UserModel.BorderStyle borderStyle, IColor borderColor);

        protected abstract IColor ConvertColor(Color color);

        public void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation)
        {
            var wb = GetWorkbook();
            _ = GetOrCreateSheet(sheetName);
            var len = colors.Length;
            var cells = GetOrCreateCells(sheetName, range);

            if (!cells.Any() || len == 0)
                return;

            var index = 0;

            ICellStyle[] styles = null;
            if (IsOpenXml)
            {
                styles = colors.Select(c =>
                {
                    var style = (XSSFCellStyle)wb.CreateCellStyle();
                    var color = (XSSFColor)ConvertColor(c);
                    style.SetFillForegroundColor(color);
                    style.FillPattern = FillPattern.SolidForeground;
                    return style;
                }).ToArray();
            }
            else
            {
                styles = colors.Select(c =>
                {
                    var style = wb.CreateCellStyle();
                    style.FillForegroundColor = ((HSSFColor)ConvertColor(c)).Indexed;
                    style.FillPattern = FillPattern.SolidForeground;
                    return style;
                }).ToArray();
            }

            if (orientation == FillOrientation.Chess)
            {
                var lastRow = cells.First().RowIndex;
                var colsCount = 0;
                foreach (var cell in cells)
                {
                    if (lastRow != cell.RowIndex)
                    {
                        if (len == colsCount || colsCount % len == 0)
                            index--;

                        lastRow = cell.RowIndex;
                        colsCount = 0;
                    }

                    if (index == len)
                        index = 0;

                    cell.CellStyle = styles[index];

                    index++;
                    colsCount++;
                }
            }
            else
            {
                var firstCell = cells.First();
                int currentCol, currentRow, firstCol;

                switch (orientation)
                {
                    case FillOrientation.Horizontal:

                        currentRow = firstCell.RowIndex;
                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.RowIndex)
                            {
                                currentRow = cell.RowIndex;
                                if (++index == len)
                                    index = 0;
                            }

                            cell.CellStyle = styles[index];
                        }

                        break;

                    case FillOrientation.Vertical:

                        firstCol = firstCell.ColumnIndex;
                        currentCol = firstCol;

                        foreach (var cell in cells)
                        {
                            if (currentCol != cell.ColumnIndex)
                            {
                                currentCol = cell.ColumnIndex;
                                if (++index == len || currentCol == firstCol)
                                    index = 0;
                            }

                            cell.CellStyle = styles[index];
                        }

                        break;
                    case FillOrientation.DiagonalForward:

                        currentCol = firstCell.ColumnIndex;
                        currentRow = firstCell.RowIndex;

                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.RowIndex)
                            {
                                currentCol++;
                                currentRow = cell.RowIndex;
                            }

                            if (currentCol == cell.ColumnIndex)
                            {
                                cell.CellStyle = styles[index];
                                if (++index == len)
                                    index = 0;
                            }
                        }

                        break;
                    case FillOrientation.DiagonalBackward:

                        currentCol = cells.Last().ColumnIndex;
                        currentRow = firstCell.RowIndex;

                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.RowIndex)
                            {
                                currentCol--;
                                currentRow = cell.RowIndex;
                            }

                            if (currentCol == cell.ColumnIndex)
                            {
                                cell.CellStyle = styles[index];
                                if (++index == len)
                                    index = 0;
                            }
                        }

                        break;
                }
            }

            RequiresSave();
        }

        public void FreezePanes(string sheetName, int cols, int rows)
        {
            var sheet = GetWorkbook().GetSheet(sheetName);
            sheet.CreateFreezePane(cols, rows);
            RequiresSave();
        }

        public string[] GetHyperlinks(string sheetName, string range)
        {
            return EnumerateHyperlinks(sheetName, range).ToArray();
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

        public DataTable ReadRange(string sheetName, string range, bool addHeaders)
        {
            var reader = GetReader();
            var dt = new DataTable();
            var addr = new Internals.RangeAddress(range);

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

        public int RemoveHyperLinks(string sheetName, string range)
        {
            int count = 0;
            foreach (var cell in GetUsedCells(sheetName, range))
            {
                if (cell.Hyperlink == null)
                    continue;

                count++;
                cell.RemoveHyperlink();
            }

            if (count > 0)
                RequiresSave();

            return count;
        }

        public void Save()
        {
            if (_requiresSave)
            {
                using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
                {
                    GetWorkbook().Write(fs);
                }
            }
        }

        public void WriteCell(string sheetName, string cell, object value)
        {
            var sheetCell = GetOrCreateCell(sheetName, cell);
            SetCellValue(sheetCell, value);
            RequiresSave();
        }

        public virtual void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders)
        {
            if (dataTable.Rows.Count == 0)
                return;

            var wb = GetWorkbook();
            var sheet = GetOrCreateSheet(sheetName);
            var cellRef = new CellReference(cell);
            var rowIndex = cellRef.Row;
            var colIndex = cellRef.Col;

            if (addHeaders)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var format = wb.CreateDataFormat();

                foreach (DataColumn col in dataTable.Columns)
                {
                    var style = wb.CreateCellStyle();
                    var sheetCell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    sheetCell.SetCellValue(col.ColumnName);

                    if (col.DataType == typeof(int))
                    {
                        style.DataFormat = format.GetFormat("0");
                    }
                    else if (col.DataType == typeof(double) || col.DataType == typeof(decimal))
                    {
                        style.DataFormat = format.GetFormat("0.00");
                    }
                    else if (col.DataType == typeof(DateTime))
                    {
                        var dateFormat = CultureInfo.CurrentCulture.DateTimeFormat;
                        var shortPattern = (dateFormat.ShortDatePattern + " " + dateFormat.ShortTimePattern).Replace("tt", "").Trim();
                        style.DataFormat = format.GetFormat(shortPattern);
                    }
                    else if (col.DataType == typeof(bool))
                    {
                        style.Alignment = HorizontalAlignment.Center;
                    }

                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    colIndex++;
                }

                colIndex = cellRef.Col;
                rowIndex++;
            }

            foreach (DataRow dr in dataTable.Rows)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                for (int i = 0; i < dr.ItemArray.Length; i++)
                {
                    var sheetCell = row.GetCell(colIndex + i) ?? row.CreateCell(colIndex + i);

                    // workaround for XSSFWorkbook, the SetDefaultColumnStyle it is not working
                    // The issue was mentioned on POI project: https://bz.apache.org/bugzilla/show_bug.cgi?id=51037
                    if (IsOpenXml)
                        sheetCell.CellStyle = sheet.GetColumnStyle(colIndex + i);

                    SetCellValue(sheetCell, dr[i]);
                }

                if (++rowIndex == MaxRows)
                    break;
            }

            foreach (DataColumn col in dataTable.Columns)
                sheet.AutoSizeColumn(col.Ordinal);

            RequiresSave();
        }

        protected static void SetCellValue(ICell cell, object value)
        {
            if (value == null)
                return;

            if (value is int || value is decimal)
                value = double.Parse(value.ToString());

            if (value is double dblValue)
                cell.SetCellValue(dblValue);
            else if (value is DateTime dateV)
                cell.SetCellValue(dateV);
            else if (value is string strV)
                cell.SetCellValue(strV);
            else if (value is bool boolV)
                cell.SetCellValue(boolV.ToString().ToUpperInvariant());
            else
                cell.SetCellValue(value?.ToString());
        }

        protected IEnumerable<ICell> GetOrCreateCells(string sheetName, string cellRange)
        {
            var sheet = GetWorkbook().GetSheet(sheetName);
            var range = CellRangeAddress.ValueOf(cellRange);

            var firstRow = range.FirstRow == -1 ? sheet.FirstRowNum : range.FirstRow;
            var lastRow = range.LastRow == -1 ? sheet.LastRowNum : range.LastRow;

            for (int i = firstRow; i <= lastRow; i++)
            {
                var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                for (int j = range.FirstColumn; j <= range.LastColumn; j++)
                {
                    var cell = row.GetCell(j) ?? row.CreateCell(j);
                    yield return cell;
                }
            }
        }

        protected ICell GetOrCreateCell(string sheetName, string cellAddress)
            => GetOrCreateCell(GetOrCreateSheet(sheetName), cellAddress);

        protected ICell GetOrCreateCell(ISheet sheet, string cellAddress)
        {
            var cellRef = new CellReference(cellAddress);
            var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
            return row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);
        }

        protected ISheet GetOrCreateSheet(string sheetName)
        {
            var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);

            if (sheet == null)
            {
                if (IsNewWorkbook)
                {
                    wb.SetSheetName(0, sheetName);
                    sheet = wb.GetSheetAt(0);
                }
                else
                {
                    sheet = wb.CreateSheet(sheetName);
                }
            }

            return sheet;
        }

        protected IEnumerable<ICell> GetUsedCells(string sheetName, string cellRange)
        {
            var sheet = GetWorkbook().GetSheet(sheetName);

            if (string.IsNullOrEmpty(cellRange))
            {
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell != null)
                            yield return cell;
                    }
                }
            }
            else
            {
                var range = CellRangeAddress.ValueOf(cellRange);

                var firstRow = range.FirstRow == -1 ? sheet.FirstRowNum : range.FirstRow;
                var lastRow = range.LastRow == -1 ? sheet.LastRowNum : range.LastRow;

                for (int i = firstRow; i <= lastRow; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    for (int j = range.FirstColumn; j <= range.LastColumn; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell != null)
                            yield return cell;
                    }
                }
            }
        }

        protected abstract IWorkbook GetWorkbook();

        protected void RequiresSave() => _requiresSave = true;

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

        private static object[][] GetDataValues(IExcelDataReader reader, Internals.RangeAddress ra, out Type[] types)
        {
            types = new Type[ra.ColsUsed()];
            var values = new object[ra.RowsUsed()][];
            int i = 0;
            while (reader.Read() && i < values.Length)
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

        private static string[] GetHeaderValues(IExcelDataReader reader, bool addHeaders, Internals.RangeAddress ra)
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

                    if (--firstRow > 0)
                        continue;

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
                ra.First.Row = rowIndex;

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

        private IEnumerable<string> EnumerateHyperlinks(string sheetName, string range)
        {
            foreach (var cell in GetUsedCells(sheetName, range))
            {
                if (cell.Hyperlink != null)
                    yield return cell.Hyperlink.Address;
            }
        }

        private IExcelDataReader GetReader()
        {
            if (_reader == null)
            {
                _reader = ExcelReaderFactory.CreateReader(WorkbookFileStream.Reset(), new ExcelReaderConfiguration { LeaveOpen = true });
            }
            else
            {
                _reader.Reset();
            }

            WorkbookFileStream.Seek(0, SeekOrigin.Begin);
            return _reader;
        }

        private void OpenOrCreate()
        {
            if (!File.Exists(FilePath))
                CreateNew();

            WorkbookFileStream = new FileStream(FilePath, FileMode.Open);
        }

        public void RenameSheet(int sheetIndex, string newName)
        {
            GetWorkbook().SetSheetName(sheetIndex, newName);
            RequiresSave();
        }

        public void DeleteSheet(string sheetName)
        {
            var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet == null)
                return;

            wb.RemoveSheetAt(wb.GetSheetIndex(sheet));
            RequiresSave();
        }

        public void MergeRange(string sheetName, string range)
        {
            GetOrCreateSheet(sheetName).AddMergedRegion(CellRangeAddress.ValueOf(range));
            RequiresSave();
        }

        public void MoveSheet(string sheetName, int index, bool makeACopy = false, string copySheetName = null)
        {
            var wb = GetWorkbook();
            var sheet = GetOrCreateSheet(sheetName);

            if (makeACopy)
            {
                sheet.CopyTo(wb, copySheetName, true, true);
                if (index > -1)
                    MoveSheet(copySheetName, index);
            }
            else
            {
                if (index < 0)
                    index = wb.NumberOfSheets + index;

                if (wb.GetSheetIndex(sheet) != index)
                    wb.SetSheetOrder(sheetName, index);
            }

            RequiresSave();
        }
    }
}