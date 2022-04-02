using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Autossential.Workbook.Core.Adapters
{
    public class OLE2WorkbookAdapter : WorkbookAdapterBase
    {
        public OLE2WorkbookAdapter(string filePath) : base(filePath)
        {
        }

        private HSSFWorkbook _workbook;
        private HSSFWorkbook GetWorkbook()
        {
            if (_workbook == null)
                _workbook = new HSSFWorkbook(WorkbookFileStream);

            return _workbook;
        }

        public override async Task<bool> AddHyperLinkAsync(string sheetName, string cellAddress, string label, string address, string tooltip)
        {
            RequiresSave();

            return await Task.Run(() =>
            {
                var cell = GetOrCreateCell(GetOrCreateSheet(sheetName), cellAddress);

                var linkType = HyperlinkType.Document;

                if (Regex.IsMatch(address, "(https?|ftp)://", RegexOptions.IgnoreCase))
                {
                    linkType = HyperlinkType.Url;
                }
                else if (address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    linkType = HyperlinkType.Email;
                }
                else if (!string.IsNullOrEmpty(Path.GetExtension(address)))
                {
                    linkType = HyperlinkType.File;
                }

                cell.Hyperlink = new HSSFHyperlink(linkType) { Address = address };
                if (string.IsNullOrEmpty(label))
                    label = address;

                cell.SetCellValue(label);
                return true;
            });
        }


        public override async Task<string[]> GetHyperlinksAsync(string sheetName, string range)
        {
            return await Task.Run(() => GetHyperlinks(sheetName, range).ToArray());
        }
        private IEnumerable<string> GetHyperlinks(string sheetName, string range)
        {
            foreach (var cell in GetCells(sheetName, range))
            {
                if (cell.Hyperlink != null)
                    yield return cell.Hyperlink.Address;
            }
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        private IEnumerable<ICell> GetCells(string sheetName, string cellRange)
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

        public override async Task<int> RemoveHyperlinksAsync(string sheetName, string range)
        {
            var result = await Task.Run(() =>
            {
                int count = 0;
                foreach (var cell in GetCells(sheetName, range))
                {
                    if (cell.Hyperlink == null)
                        continue;

                    count++;
                    cell.RemoveHyperlink();
                }
                return count;
            });

            if (result > 0)
                RequiresSave();

            return result;
        }

        public override Action GetSaveHandler()
        {
            return () =>
            {
                using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
                {
                    GetWorkbook().Write(fs);
                }
            };
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

        public override async Task WriteRangeAsync(string sheetName, string cellAddress, DataTable value, bool addHeaders)
        {
            if (value.Rows.Count == 0)
                return;

            RequiresSave();

            await Task.Run(() =>
            {
                var sheet = GetOrCreateSheet(sheetName);
                var cellRef = new CellReference(cellAddress);
                var rowIndex = cellRef.Row;
                var colIndex = cellRef.Col;

                if (addHeaders)
                {
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    foreach (DataColumn col in value.Columns)
                    {
                        var cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                        cell.SetCellValue(col.ColumnName);
                        colIndex++;
                    }

                    colIndex = cellRef.Col;
                    rowIndex++;
                }

                foreach (DataRow dr in value.Rows)
                {
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    for (int i = 0; i < dr.ItemArray.Length; i++)
                    {
                        var cell = row.GetCell(colIndex + i) ?? row.CreateCell(colIndex + i);
                        SetCellValue(cell, dr[i]);
                    }
                    rowIndex++;
                }
            });
        }

        private void SetCellValue(ICell cell, object value)
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
            else
                cell.SetCellValue(value?.ToString());
        }

        public override async Task WriteCellAsync(string sheetName, string cellAddress, object value)
        {
            RequiresSave();
            await Task.Run(() =>
            {
                var cell = GetOrCreateCell(GetOrCreateSheet(sheetName), cellAddress);
                SetCellValue(cell, value);
            });
        }

        private ICell GetOrCreateCell(ISheet sheet, string cellAddress)
        {
            var cellRef = new CellReference(cellAddress);
            var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
            return row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);
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
    }
}