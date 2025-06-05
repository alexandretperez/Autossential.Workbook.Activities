using Autossential.Workbook.Core.Internals;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using Sylvan.Data.Excel;
using System;
using System.IO;

namespace Autossential.Workbook.Core.Processors
{
    public class BIFF8WorkbookProcessor : WorkbookProcessorBase
    {
        public BIFF8WorkbookProcessor(string filePath) : base(filePath)
        {
        }

        private IWorkbook GetWorkbook()
        {
            if (_workbook == null)
            {
                WorkbookStream.Reset();
                _workbook = new HSSFWorkbook(new PushbackInputStream(new ByteArrayInputStream(WorkbookStream.ToArray())));
            }

            return _workbook;
        }

        private IWorkbook _workbook;
        public override void Save()
        {
            if (!RequiresSave)
                return;

            WorkbookStream.SetLength(0);
            GetWorkbook().Write(WorkbookStream);

            var bytes = WorkbookStream.ToArray();
            File.WriteAllBytes(FilePath, bytes);
        }

        public override void RenameSheet(int sheetIndex, string newSheetName)
        {
            var wb = GetWorkbook();

            if (sheetIndex < 0 || sheetIndex >= wb.NumberOfSheets)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, "Sheet index is out of range");

            wb.SetSheetName(sheetIndex, newSheetName);
            wb.Write(WorkbookStream);

            RequiresSave = true;
        }

        public override void Dispose(bool disposing)
        {
            _workbook?.Close();
        }

        public override void RenameSheet(string fromSheetName, string toSheetName)
        {
            var index = Array.IndexOf(GetSheetNames(), fromSheetName);
            RenameSheet(index, toSheetName);
        }

        protected override ExcelDataReader GetReader(ExcelDataReaderOptions options = null) =>
            ExcelDataReader.Create(WorkbookStream.Reset(), ExcelWorkbookType.Excel, options);

        public override void DeleteSheet(string sheetName)
        {
            var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet == null)
                return;

            wb.RemoveSheetAt(wb.GetSheetIndex(sheet));
            wb.Write(WorkbookStream);

            RequiresSave = true;
        }

        public override void ActivateSheet(string sheetName)
        {
            var index = GetWorkbook().GetSheetIndex(sheetName);
            if (index == -1)
                throw new ArgumentException("Sheet not found", nameof(sheetName));

            ActivateSheet(index);
        }

        public override void ActivateSheet(int sheetIndex)
        {
            var wb = GetWorkbook();
            if (sheetIndex < 0 || sheetIndex > wb.NumberOfSheets)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, "Sheet index is out of range");

            wb.SetActiveSheet(sheetIndex);
            wb.Write(WorkbookStream);

            RequiresSave = true;
        }

        public override (int index, string name) GetActiveSheet()
        {
            var wb = GetWorkbook();
            if (wb.ActiveSheetIndex >= 0)
                return (wb.ActiveSheetIndex, wb.GetSheetName(wb.ActiveSheetIndex));

            return (-1, null);
        }

        internal override RangeReference ResolveRange(string range)
        {
            if (string.IsNullOrEmpty(range))
                throw new ArgumentException("Range cannot be null or empty", nameof(range));

            var rangeRef = new BIFF8RangeReference(range);
            if (!rangeRef.IsValid)
                throw new ArgumentException("Invalid range format", nameof(range));

            return rangeRef;
        }
    }
}