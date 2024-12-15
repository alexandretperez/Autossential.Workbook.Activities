using Autossential.Workbook.Core.Internals;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Processors
{
    public class OpenXMLWorkbookProcessor : WorkbookProcessorBase
    {
        private SpreadsheetDocument _document;

        public OpenXMLWorkbookProcessor(string filePath) : base(filePath)
        {
            
        }

        public override void RenameSheet(int sheetIndex, string newSheetName)
        {
            var doc = GetDocument();
            var wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Descendants<Sheet>();
            if (sheetIndex < 0 || sheetIndex >= sheets.Count())
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, "Sheet index is out of range");

            var sheet = sheets.ElementAt(sheetIndex);
            sheet.Name = newSheetName;

            SaveInMemory();

            RequiresSave = true;
        }

        private void SaveInMemory()
        {
            _document.Save();
            _document.Close();
            _document = null;
        }

        public override void Save()
        {
            if (!RequiresSave)
                return;

            var bytes = WorkbookStream.ToArray();
            File.WriteAllBytes(FilePath, bytes);
        }

        public override void Dispose(bool disposing)
        {
            if (disposing)
                _document?.Dispose();

            base.Dispose(disposing);
        }

        private SpreadsheetDocument GetDocument() => _document ??= SpreadsheetDocument.Open(WorkbookStream.Reset(), true, new OpenSettings { AutoSave = false });

        public override void RenameSheet(string fromSheetName, string toSheetName)
        {
            var index = Array.IndexOf(GetSheetNames(), fromSheetName);
            RenameSheet(index, toSheetName);
        }
    }
}