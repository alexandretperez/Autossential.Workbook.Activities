using Autossential.Workbook.Core.Internals;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
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
            _workbook.Write(WorkbookStream);

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
    }
}