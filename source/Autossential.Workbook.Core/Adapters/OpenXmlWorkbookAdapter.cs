using Autossential.Workbook.Core.Internals;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Adapters
{
    public class OpenXmlWorkbookAdapter : WorkbookAdapterBase
    {
        private XSSFWorkbook _workbook;

        public OpenXmlWorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override int MaxRows => 1_048_576;

        public override bool IsOpenXml => true;

        public override void CreateNew()
        {
            using (var fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                workbook.CreateSheet("Sheet1");
                workbook.Write(fs);
            }
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        protected override IColor ConvertColor(Color color)
        {
            return new XSSFColor(color);
        }

        protected override IWorkbook GetWorkbook()
        {
            if (_workbook == null)
            {
                try
                {
                    _workbook = new XSSFWorkbook(WorkbookFileStream.Reset());
                }
                catch (Exception e)
                {
                    Trace.WriteLine(e.Message);
                }
            }

            return _workbook;
        }

        protected override void ApplyBorderStyle(ICellStyle cellStyle, string anchors, BorderStyle borderStyle, IColor borderColor)
        {
            var style = (XSSFCellStyle)cellStyle;
            var color = (XSSFColor)borderColor;
            if (anchors.Contains('T'))
            {
                style.BorderTop = borderStyle;
                style.SetTopBorderColor(color);
            }
            if (anchors.Contains('B'))
            {
                style.BorderBottom = borderStyle;
                style.SetBottomBorderColor(color);
            }
            if (anchors.Contains('L'))
            {
                style.BorderLeft = borderStyle;
                style.SetLeftBorderColor(color);
            }
            if (anchors.Contains('R'))
            {
                style.BorderRight = borderStyle;
                style.SetRightBorderColor(color);
            }
        }

        protected override void ApplyBackgroundStyle(ICellStyle cellStyle, Color color)
        {
            var style = (XSSFCellStyle)cellStyle;
            style.SetFillForegroundColor((XSSFColor)ConvertColor(color));
            style.FillPattern = FillPattern.SolidForeground;
        }
    }
}