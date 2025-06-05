using Autossential.Workbook.Core.Internals;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Sylvan.Data.Excel;
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
            _document.WorkbookPart.Workbook.Save();
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

        protected override ExcelDataReader GetReader(ExcelDataReaderOptions options = null) =>
            ExcelDataReader.Create(WorkbookStream.Reset(), ExcelWorkbookType.ExcelXml, options);

        public override void DeleteSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public override void ActivateSheet(string sheetName)
        {
            var document = GetDocument();
            var wbPart = document.WorkbookPart;
            var wb = wbPart.Workbook;
            var sheets = wb.Sheets;
            int sheetIndex = sheets.Elements<Sheet>()
                                   .ToList()
                                   .FindIndex(sheet => sheet.Name == sheetName);

            if (sheetIndex == -1)
                throw new ArgumentException("Sheet not found", nameof(sheetName));

            ActivateSheet(sheetIndex);
        }

        //public override void ActivateSheet(int sheetIndex)
        //{
        //    var document = GetDocument();
        //    var wbPart = document.WorkbookPart;
        //    var wb = wbPart.Workbook;
        //    var sheets = wb.Sheets;
        //    if (sheetIndex < 0 || sheetIndex >= sheets.Count())
        //        throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, "Sheet index is out of range");

        //    var workbookView = wb.BookViews?.OfType<WorkbookView>().FirstOrDefault();
        //    if (workbookView == null)
        //    {
        //        workbookView = new WorkbookView();
        //        wb.BookViews ??= new BookViews();
        //        wb.BookViews.Append(workbookView);
        //    }

        //    workbookView.ActiveTab = (uint)sheetIndex;
        //    SaveInMemory();
        //    RequiresSave = true;
        //}

        public override void ActivateSheet(int sheetIndex)
        {
            var doc = GetDocument();
            WorkbookPart workbookPart = doc.WorkbookPart;
            Sheets sheets = workbookPart.Workbook.Sheets;
            Sheet sheetToActivate = sheets.Elements<Sheet>().ElementAtOrDefault(sheetIndex);

            if (sheetToActivate != null)
            {
                // Define a aba ativa corretamente
                WorkbookView workbookView = workbookPart.Workbook.BookViews.Elements<WorkbookView>().FirstOrDefault();
                if (workbookView != null)
                {
                    workbookView.ActiveTab = (uint)sheetIndex;
                }

                // Percorre todas as planilhas e remove a seleção de qualquer outra aba
                foreach (Sheet sheet in sheets.Elements<Sheet>())
                {
                    WorksheetPart sheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetView sheetView = sheetPart.Worksheet.Elements<SheetViews>().FirstOrDefault()?.Elements<SheetView>().FirstOrDefault();

                    if (sheetView != null)
                    {
                        sheetView.TabSelected = (sheet == sheetToActivate); // Apenas a planilha selecionada será marcada
                    }
                }

                SaveInMemory();
                RequiresSave = true;
            }
        }



        public override (int index, string name) GetActiveSheet()
        {
            var document = GetDocument();
            WorkbookPart workbookPart = document.WorkbookPart;
            var workbook = workbookPart.Workbook;
            var workbookView = workbook.BookViews?.OfType<WorkbookView>().FirstOrDefault();
            uint activeIndex = workbookView?.ActiveTab ?? 0;
            Sheets sheets = workbook.Sheets;
            if (sheets.ElementAt((int)activeIndex) is Sheet activeSheet)
                return ((int)activeIndex, activeSheet.Name);

            return (-1, null);
        }

        internal override RangeReference ResolveRange(string range)
        {
            if (string.IsNullOrEmpty(range))
                throw new ArgumentException("Range cannot be null or empty", nameof(range));

            var rangeRef = new OpenXMLRangeReference(range);
            if (!rangeRef.IsValid)
                throw new ArgumentException("Invalid range format", nameof(range));

            return rangeRef;
        }
    }
}