using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
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
                _workbook = new HSSFWorkbook(File.OpenRead(FilePath));

            return _workbook;
        }

        public override async Task<bool> AddHyperLinkAsync(string sheetName, string cellAddress, string label, string address, string tooltip)
        {
            RequiresSave();

            return await Task.Run(() =>
            {
                var sheet = GetWorkbook().GetSheet(sheetName);
                var cellRef = new CellReference(cellAddress);
                var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
                var cell = row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);

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
            throw new NotImplementedException();
        }
    }
}
