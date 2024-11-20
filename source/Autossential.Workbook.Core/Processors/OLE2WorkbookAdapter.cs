using Autossential.Workbook.Core.Internals;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Adapters
{
    public class OLE2WorkbookAdapter : WorkbookAdapterBase
    {
        public OLE2WorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override int MaxRows => 65_536;

        public override bool IsOpenXml => false;

        private HSSFWorkbook _workbook;

        protected override IWorkbook GetWorkbook()
        {
            if (_workbook == null)
            {
                try
                {
                    _workbook = new HSSFWorkbook(WorkbookFileStream.Reset());
                }
                catch (Exception e)
                {
                    Trace.WriteLine(e.Message);
                }
            }

            return _workbook;
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

        public void WriteRangeX(string sheetName, string cell, DataTable dataTable, bool addHeaders)
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
                foreach (DataColumn col in dataTable.Columns)
                {
                    var sheetCell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    sheetCell.SetCellValue(col.ColumnName);

                    if (col.DataType == typeof(int))
                    {
                        var style = wb.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(double) || col.DataType == typeof(decimal))
                    {
                        var style = wb.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(DateTime))
                    {
                        var dateFormat = CultureInfo.CurrentCulture.DateTimeFormat;
                        var shortPattern = (dateFormat.ShortDatePattern + " " + dateFormat.ShortTimePattern).Replace("tt", "").Trim();
                        var style = wb.CreateCellStyle();
                        style.DataFormat = wb.CreateDataFormat().GetFormat(shortPattern);
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(bool))
                    {
                        var style = wb.CreateCellStyle();
                        style.Alignment = HorizontalAlignment.Center;
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else
                    {
                        var style = wb.CreateCellStyle();
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

                if (++rowIndex == MaxRows)
                    break;
            }

            foreach (DataColumn col in dataTable.Columns)
                sheet.AutoSizeColumn(col.Ordinal);

            RequiresSave();
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        private HSSFPalette _palette;
        protected override IColor ConvertColor(Color color)
        {
            _palette = _palette ?? ((HSSFWorkbook)GetWorkbook()).GetCustomPalette();
            return _palette.FindSimilarColor(color.R, color.G, color.B);
        }

        protected override void ApplyBorderStyle(ICellStyle cellStyle, string anchors, BorderStyle borderStyle, IColor borderColor)
        {
            var style = (HSSFCellStyle)cellStyle;
            var color = (HSSFColor)borderColor;
            if (anchors.Contains('T'))
            {
                style.BorderTop = borderStyle;
                style.TopBorderColor = color.Indexed;
            }
            if (anchors.Contains('B'))
            {
                style.BorderBottom = borderStyle;
                style.BottomBorderColor = color.Indexed;
            }
            if (anchors.Contains('L'))
            {
                style.BorderLeft = borderStyle;
                style.LeftBorderColor = color.Indexed;
            }
            if (anchors.Contains('R'))
            {
                style.BorderRight = borderStyle;
                style.RightBorderColor = color.Indexed;
            }
        }

        public override void DrawBorder(string sheetName, string range, Enums.Border border, Enums.BorderStyle style, Color color)
        {
            base.DrawBorder(sheetName, range, border, style, color);
            OptimizeStyles();
        }

        public override void FillColor(string sheetName, string range, Color[] colors, Enums.FillPattern orientation)
        {
            base.FillColor(sheetName, range, colors, orientation);
            OptimizeStyles();
        }

        private void OptimizeStyles()
        {
            HSSFOptimiser.OptimiseCellStyles((HSSFWorkbook)GetWorkbook());
        }

        protected override void ApplyBackgroundStyle(ICellStyle cellStyle, Color color)
        {
            cellStyle.FillForegroundColor = ((HSSFColor)ConvertColor(color)).Indexed;
            cellStyle.FillPattern = FillPattern.SolidForeground;
        }
    }
}