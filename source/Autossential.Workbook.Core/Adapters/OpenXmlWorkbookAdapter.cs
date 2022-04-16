using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
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
            foreach (var cell in GetCells(sheetName, range))
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

            RequiresSave();
        }

        private IEnumerable<string> EnumerateHyperlinks(string sheetName, string range)
        {
            foreach (var cell in GetCells(sheetName, range))
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

        private IXLCells GetCells(string sheetName, string range)
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
    }
}