using Autossential.Workbook.Core.Enums;
using Autossential.Workbook.Core.Internals;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Autossential.Workbook.Core.Adapters
{
    public class OLE2WorkbookAdapter : WorkbookAdapterBase
    {
        public const int MAX_ROWS = 65536;

        private HSSFWorkbook _workbook;

        public OLE2WorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override void AddHyperLink(string sheetName, string cell, string label, string link, string tooltip)
        {
            var sheetCell = GetOrCreateCell(GetOrCreateSheet(sheetName), cell);
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

            sheetCell.Hyperlink = new HSSFHyperlink(linkType) { Address = link };
            if (string.IsNullOrEmpty(label))
                label = link;

            sheetCell.SetCellValue(label);

            RequiresSave();
        }

        public override void AppendRange(string sheetName, DataTable dataTable)
        {
            var sheet = GetOrCreateSheet(sheetName);

            var cell = "A1";
            var row = sheet.GetRow(sheet.LastRowNum);
            if (row != null)
                cell = new CellReference(row.RowNum + 1, row.FirstCellNum).FormatAsString();

            WriteRange(sheetName, cell, dataTable, false);
        }

        public override void CreateNew()
        {
            using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
            {
                var workbook = new HSSFWorkbook();
                workbook.CreateSheet("Sheet1");
                workbook.Write(fs);
            }
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        public override void FreezePanes(string sheetName, int cols, int rows)
        {
            var sheet = GetWorkbook().GetSheet(sheetName);
            sheet.CreateFreezePane(cols, rows);
            RequiresSave();
        }

        public override string[] GetHyperlinks(string sheetName, string range)
            => EnumerateLinks(sheetName, range).ToArray();

        public override Action GetSaveHandler() => () =>
                {
                    using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
                    {
                        GetWorkbook().Write(fs);
                    }
                };

        public override int RemoveHyperLinks(string sheetName, string range)
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

        public override void WriteCell(string sheetName, string cell, object value)
        {
            var sheetCell = GetOrCreateCell(GetOrCreateSheet(sheetName), cell);
            SetCellValue(sheetCell, value);
        }

        public override void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders)
        {
            if (dataTable.Rows.Count == 0)
                return;

            var sheet = GetOrCreateSheet(sheetName);
            var cellRef = new CellReference(cell);
            var rowIndex = cellRef.Row;
            var colIndex = cellRef.Col;

            if (addHeaders)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                foreach (DataColumn col in dataTable.Columns)
                {
                    var sheetCell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    sheetCell.SetCellValue(col.ColumnName);

                    if (col.DataType == typeof(int))
                    {
                        var style = _workbook.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(double) || col.DataType == typeof(decimal))
                    {
                        var style = _workbook.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(DateTime))
                    {
                        var dateFormat = CultureInfo.CurrentCulture.DateTimeFormat;
                        var shortPattern = (dateFormat.ShortDatePattern + " " + dateFormat.ShortTimePattern).Replace("tt", "").Trim();
                        var style = _workbook.CreateCellStyle();
                        style.DataFormat = _workbook.CreateDataFormat().GetFormat(shortPattern);
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(bool))
                    {
                        var style = _workbook.CreateCellStyle();
                        style.Alignment = HorizontalAlignment.Center;
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else
                    {
                        var style = _workbook.CreateCellStyle();
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }

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
                    SetCellValue(sheetCell, dr[i]);
                }

                if (++rowIndex == MAX_ROWS)
                    break;
            }

            foreach (DataColumn col in dataTable.Columns)
                sheet.AutoSizeColumn(col.Ordinal);

            RequiresSave();
        }

        private static ICell GetOrCreateCell(ISheet sheet, string cellAddress)
        {
            var cellRef = new CellReference(cellAddress);
            var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
            return row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);
        }

        private static void SetCellValue(ICell cell, object value)
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

        private IEnumerable<string> EnumerateLinks(string sheetName, string range)
        {
            foreach (var cell in GetUsedCells(sheetName, range))
            {
                if (cell.Hyperlink != null)
                    yield return cell.Hyperlink.Address;
            }
        }

        private IEnumerable<ICell> GetUsedCells(string sheetName, string cellRange)
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

        private IEnumerable<ICell> GetCells(string sheetName, string cellRange)
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

        private ISheet GetOrCreateSheet(string sheetName)
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
                    sheet = GetWorkbook().CreateSheet(sheetName);
                }
            }

            return sheet;
        }

        private HSSFWorkbook GetWorkbook()
        {
            if (_workbook == null)
                _workbook = new HSSFWorkbook(WorkbookFileStream);

            return _workbook;
        }


        private NPOI.HSSF.Util.HSSFColor GetDarkerColor(NPOI.HSSF.Util.HSSFColor xlsColor, HSSFPalette palette)
        {
            var rgb = xlsColor.RGB;
            while (rgb[0] + rgb[1] + rgb[2] > 0)
            {
                if (rgb[0] > 0) rgb[0]--;
                if (rgb[1] > 0) rgb[1]--;
                if (rgb[2] > 0) rgb[2]--;

                var dark = palette.FindSimilarColor(rgb[0], rgb[1], rgb[2]);
                if (!dark.RGB.SequenceEqual(xlsColor.RGB))
                    return dark;
            }
            return xlsColor;
        }

        public override void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
        {
            throw new NotImplementedException();
        }

        public override void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation)
        {
            var wb = GetWorkbook();
            var sheet = GetOrCreateSheet(sheetName);
            var palette = wb.GetCustomPalette();
            var len = colors.Length;
            var cells = GetCells(sheetName, range);

            if (!cells.Any() || len == 0)
                return;

            var index = 0;

            var styles = colors.Select(c =>
            {
                var style = wb.CreateCellStyle();
                style.FillForegroundColor = palette.FindSimilarColor(c.R, c.G, c.B).Indexed;
                style.FillPattern = FillPattern.SolidForeground;
                return style;
            }).ToArray();

            if (orientation == FillOrientation.Chess)
            {
                foreach (var cell in cells)
                {
                    cell.CellStyle = styles[index];
                    if (++index == len)
                        index = 0;
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
                    case FillOrientation.DiagonalLeft:

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
                    case FillOrientation.DiagonalRight:

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
    }
}