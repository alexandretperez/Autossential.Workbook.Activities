using Autossential.Workbook.Core.Enums;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace Autossential.Workbook.Core.Adapters
{
    public class OpenXmlWorkbookAdapter : WorkbookAdapterBase
    {
        private XLWorkbook _workbook = null;

        public OpenXmlWorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override void AddHyperLink(string sheetName, string cell, string label, string link, string tooltip)
        {
            var sheetCell = GetOrCreateSheet(sheetName).Cell(cell);
            sheetCell.Hyperlink = new XLHyperlink(link, tooltip);
            if (string.IsNullOrEmpty(label))
                label = link;

            sheetCell.SetValue(label);
            RequiresSave();
        }

        public override void CreateNew()
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Sheet1");
                workbook.SaveAs(FilePath);
            }
        }

        public override void Dispose(bool disposing)
        {
            if (!disposing && _workbook != null)
                _workbook.Dispose();
        }

        public override void FreezePanes(string sheetName, int cols, int rows)
        {
            var sheet = GetWorkbook().Worksheet(sheetName);
            sheet.SheetView.Freeze(rows, cols);
            RequiresSave();
        }

        public override string[] GetHyperlinks(string sheetName, string range)
            => EnumerateHyperlinks(sheetName, range).ToArray();

        public override Action GetSaveHandler() => GetWorkbook().Save;

        public override int RemoveHyperLinks(string sheetName, string range)
        {
            var count = 0;
            foreach (var cell in GetUsedCells(sheetName, range))
            {
                if (cell.HasHyperlink)
                {
                    cell.Hyperlink.Delete();
                    count++;
                }
            }

            if (count > 0)
                RequiresSave();

            return count;
        }

        public override void WriteCell(string sheetName, string cell, object value)
        {
            var sheetCell = GetOrCreateSheet(sheetName).Cell(cell);
            sheetCell.Value = value;
            RequiresSave();
        }

        public override void AppendRange(string sheetName, DataTable dataTable)
        {
            var sheet = GetOrCreateSheet(sheetName);
            var cell = "A1";
            var rangeUsed = sheet.RangeUsed();
            if (rangeUsed != null)
                cell = rangeUsed.LastRow().RowBelow(1).RangeAddress.FirstAddress.ToString();

            WriteRange(sheetName, cell, dataTable, false);
        }

        public override void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders)
        {
            if (dataTable.Rows.Count == 0)
                return;

            var sheet = GetOrCreateSheet(sheetName);
            var sheetCell = sheet.Cell(cell);
            var rowIndex = sheetCell.Address.RowNumber;
            var colIndex = sheetCell.Address.ColumnNumber;

            if (addHeaders)
            {
                foreach (DataColumn col in dataTable.Columns)
                    sheet.Cell(rowIndex, colIndex++).SetValue(col.ColumnName);

                colIndex = sheetCell.Address.ColumnNumber;
                rowIndex++;
            }

            foreach (DataRow dr in dataTable.Rows)
            {
                for (int i = 0; i < dr.ItemArray.Length; i++)
                    sheet.Cell(rowIndex, colIndex + i).SetValue(dr[i]);

                rowIndex++;
            }

            for (int i = sheetCell.Address.ColumnNumber; i <= dataTable.Columns.Count; i++)
                sheet.Column(i).AdjustToContents();

            RequiresSave();
        }

        private IEnumerable<string> EnumerateHyperlinks(string sheetName, string range)
        {
            foreach (var cell in GetUsedCells(sheetName, range))
            {
                if (cell.HasHyperlink)
                {
                    if (cell.Hyperlink.IsExternal)
                    {
                        yield return cell.Hyperlink.ExternalAddress.ToString();
                        continue;
                    }

                    yield return cell.Hyperlink.InternalAddress.ToString();
                }
            }
        }

        private IXLCells GetUsedCells(string sheetName, string range)
        {
            var sheet = GetWorkbook().Worksheet(sheetName);
            return string.IsNullOrEmpty(range) ? sheet.CellsUsed() : sheet.Range(range).CellsUsed();
        }

        private IXLWorksheet GetOrCreateSheet(string sheetName)
        {
            var wb = GetWorkbook();
            if (!wb.TryGetWorksheet(sheetName, out var sheet))
            {
                if (IsNewWorkbook)
                {
                    sheet = wb.Worksheet(1);
                    sheet.Name = sheetName;
                }
                else
                {
                    sheet = wb.AddWorksheet(sheetName);
                }
            }

            return sheet;
        }

        private XLWorkbook GetWorkbook()
        {
            if (_workbook == null)
                _workbook = new XLWorkbook(WorkbookFileStream, XLEventTracking.Disabled);

            return _workbook;
        }

        public override void DrawBorder(string sheetName, string range, Border border, BorderStyle style, Color color)
        {
            throw new NotImplementedException();
        }

        public override void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation)
        {
            var cells = GetOrCreateSheet(sheetName).Cells(range);
            var len = colors.Length;

            if (!cells.Any() || len == 0)
                return;

            var index = 0;

            if (orientation == FillOrientation.Chess)
            {
                foreach (var cell in cells)
                {
                    cell.Style.Fill.BackgroundColor = XLColor.FromColor(colors[index]);
                    if (++index == len)
                        index = 0;
                }
            }
            else
            {
                var firstCellAddr = cells.First().Address;
                int currentCol, currentRow, firstCol;

                switch (orientation)
                {
                    case FillOrientation.Horizontal:

                        currentRow = firstCellAddr.RowNumber;

                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.Address.RowNumber)
                            {
                                currentRow = cell.Address.RowNumber;
                                if (++index == len)
                                    index = 0;
                            }

                            cell.Style.Fill.BackgroundColor = XLColor.FromColor(colors[index]);
                        }

                        break;

                    case FillOrientation.Vertical:

                        firstCol = firstCellAddr.ColumnNumber;
                        currentCol = firstCol;

                        foreach (var cell in cells)
                        {
                            if (currentCol != cell.Address.ColumnNumber)
                            {
                                currentCol = cell.Address.ColumnNumber;
                                if (++index == len || currentCol == firstCol)
                                    index = 0;
                            }

                            cell.Style.Fill.BackgroundColor = XLColor.FromColor(colors[index]);
                        }

                        break;

                    case FillOrientation.DiagonalLeft:

                        currentCol = firstCellAddr.ColumnNumber;
                        currentRow = firstCellAddr.RowNumber;

                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.Address.RowNumber)
                            {
                                currentCol++;
                                currentRow = cell.Address.RowNumber;
                            }

                            if (currentCol == cell.Address.ColumnNumber)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromColor(colors[index]);
                                if (++index == len)
                                    index = 0;
                            }
                        }

                        break;

                    case FillOrientation.DiagonalRight:

                        currentCol = cells.Last().Address.ColumnNumber;
                        currentRow = firstCellAddr.RowNumber;

                        foreach (var cell in cells)
                        {
                            if (currentRow != cell.Address.RowNumber)
                            {
                                currentCol--;
                                currentRow = cell.Address.RowNumber;
                            }

                            if (currentCol == cell.Address.ColumnNumber)
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.FromColor(colors[index]);
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