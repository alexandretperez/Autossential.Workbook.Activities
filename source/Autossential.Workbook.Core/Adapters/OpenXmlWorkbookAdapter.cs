using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Autossential.Workbook.Core.Adapters
{
    public class OpenXmlWorkbookAdapter : WorkbookAdapterBase
    {
        private XLWorkbook _workbook = null;

        private XLWorkbook GetWorkbook()
        {
            if (_workbook == null)
                _workbook = new XLWorkbook(WorkbookFileStream, XLEventTracking.Disabled);

            return _workbook;
        }

        public OpenXmlWorkbookAdapter(string filePath) : base(filePath)
        {
        }

        private IXLCells GetCells(string sheetName, string range)
        {
            var sheet = GetWorkbook().Worksheet(sheetName);
            return string.IsNullOrEmpty(range) ? sheet.CellsUsed() : sheet.Range(range).CellsUsed();
        }

        public override void Dispose(bool disposing)
        {
            if (!disposing && _workbook != null)
                _workbook.Dispose();
        }

        public override async Task<bool> AddHyperLinkAsync(string sheetName, string cellAddress, string label, string address, string tooltip)
        {
            RequiresSave();
            return await Task.Run(() =>
            {
                var cell = GetOrCreateSheet(sheetName).Cell(cellAddress);
                cell.Hyperlink = new XLHyperlink(address, tooltip);
                if (string.IsNullOrEmpty(label))
                    label = address;

                cell.SetValue(label);
                return true;
            });
        }

        public override async Task<int> RemoveHyperlinksAsync(string sheetName, string range)
        {
            var result = await Task.Run(() =>
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

                return count;
            });

            if (result > 0)
                RequiresSave();

            return result;
        }

        public override async Task<string[]> GetHyperlinksAsync(string sheetName, string range)
        {
            return await Task.Run(() => GetHyperlinks(sheetName, range).ToArray());
        }

        private IEnumerable<string> GetHyperlinks(string sheetName, string range)
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

        public override Action GetSaveHandler() => GetWorkbook().Save;

        public override void CreateNew()
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Sheet1");
                workbook.SaveAs(FilePath);
            }
        }

        public override async Task WriteRangeAsync(string sheetName, string cellAddress, DataTable value, bool addHeaders)
        {
            if (value.Rows.Count == 0)
                return;

            RequiresSave();

            await Task.Run(() =>
            {
                var sheet = GetOrCreateSheet(sheetName);
                var cell = sheet.Cell(cellAddress);
                var rowIndex = cell.Address.RowNumber;
                var colIndex = cell.Address.ColumnNumber;

                if (addHeaders)
                {
                    foreach (DataColumn col in value.Columns)
                        sheet.Cell(rowIndex, colIndex++).SetValue(col.ColumnName);

                    colIndex = cell.Address.ColumnNumber;
                    rowIndex++;
                }

                foreach (DataRow dr in value.Rows)
                {
                    for (int i = 0; i < dr.ItemArray.Length; i++)
                        sheet.Cell(rowIndex, colIndex + i).SetValue(dr[i]);
                    
                    rowIndex++;
                }
            });
        }

        public override async Task WriteCellAsync(string sheetName, string cellAddress, object value)
        {
            RequiresSave();
            await Task.Run(() =>
            {
                var cell = GetOrCreateSheet(sheetName).Cell(cellAddress);
                cell.Value = value;
            });
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
    }
}