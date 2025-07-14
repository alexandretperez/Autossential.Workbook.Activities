using Autossential.Workbook.Core.Extensions;
using Autossential.Workbook.Core.Internals;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Globalization;
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
            var sheets = wbPart.Workbook.Descendants<Sheet>().ToArray();
            if (sheetIndex < 0 || sheetIndex >= sheets.Length)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, "Sheet index is out of range");

            var sheet = sheets[sheetIndex];
            if (sheet.Name == newSheetName)
                return;

            if (Array.Find(sheets, s => s.Name.Value.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)) != null)
                throw new InvalidOperationException($"The sheet name '{newSheetName}' already exists in the workbook");

            sheet.Name = newSheetName;
            SaveInMemory();
        }

        private void SaveInMemory()
        {
            RequiresSave = true;

            _document.WorkbookPart.Workbook.Save();
            _document.Save();
        }

        public override void Save()
        {
            SaveInMemory();
            _document.Close();
            _document = null;

            if (!RequiresSave)
                return;

            WorkbookStream.Position = 0;

            using var fs = File.Create(FilePath);
            WorkbookStream.CopyTo(fs, WorkbookStream.CalculateBufferSize());
        }

        public override void Dispose(bool disposing)
        {
            if (disposing)
                _document?.Dispose();

            base.Dispose(disposing);
        }

        private SpreadsheetDocument GetDocument()
        {
            if (WorkbookStream.Length == 0)
            {
                _document = SpreadsheetDocument.Create(WorkbookStream, SpreadsheetDocumentType.Workbook, false);
            }

            return _document ??= SpreadsheetDocument.Open(WorkbookStream.Reset(), true, new OpenSettings { AutoSave = false });
        }

        public override void RenameSheet(string fromSheetName, string toSheetName)
        {
            var index = Array.IndexOf(GetSheetNames(), fromSheetName);
            RenameSheet(index, toSheetName);
        }

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

            var rangeRef = RangeReference.CreateForOpenXml(range);
            if (!rangeRef.IsValid)
                throw new ArgumentException("Invalid range format", nameof(range));

            return rangeRef;
        }

        public override void WriteRange(DataTable dt, string sheetName, string startCell, bool addHeaders)
        {
            if (dt == null)
                throw new ArgumentNullException(nameof(dt), "DataTable cannot be null");

            ValidateSheetName(sheetName);

            var cellRef = CellReference.CreateForOpenXml(startCell);

            WorkbookPart wbPart;

            var doc = GetDocument();
            if (WorkbookStream.Length == 0) // New File
            {
                wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                wbPart.Workbook.AppendChild(new Sheets());
                var sheet = wbPart.CreateSheet(sheetName);
                var wsPart = wbPart.AddNewPart<WorksheetPart>();
                sheet.Id = wbPart.GetIdOfPart(wsPart);
                WriteDataWithOpenXmlWriter(dt, wsPart, cellRef.Row, cellRef.Col, true);
            }
            else
            {
                wbPart = doc.WorkbookPart;
                var sheet = wbPart.GetOrCreateSheet(sheetName);
                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
                WriteDataWithDOM(dt, wsPart, cellRef.Row, cellRef.Col, addHeaders);
            }

            SaveInMemory();
        }

        private static void RemoveCellsInRange(SheetData sheetData, uint startRow, uint endRow, uint startCol, uint endCol)
        {
            var rowsInRange = sheetData.Elements<Row>()
                .Where(r => r.RowIndex >= startRow && r.RowIndex <= endRow)
                .ToList();

            foreach (Row row in rowsInRange)
            {
                var cellsToRemove = row.Elements<Cell>().Where(c =>
                    c.CellReference != null
                    && CellReference.CreateForOpenXml(c.CellReference.Value).IsColInRange(startCol, endCol)
                ).ToArray();

                foreach (Cell cell in cellsToRemove)
                    cell.Remove();
            }
        }


        private static void WriteDataWithDOM(
            DataTable dt,
            WorksheetPart worksheetPart,
            uint startRowIndex,
            uint startColIndex,
            bool addHeaders)
        {

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            uint currentRowIndex = startRowIndex;

            // Calcular intervalo de escrita
            uint totalRows = (uint)dt.Rows.Count + (addHeaders ? 1u : 0u);
            uint endRowIndex = startRowIndex + totalRows - 1;
            uint endColIndex = startColIndex + (uint)dt.Columns.Count - 1;

            // Remover células existentes no intervalo
            RemoveCellsInRange(sheetData, startRowIndex, endRowIndex, startColIndex, endColIndex);

            // Escrever cabeçalhos se necessário
            if (addHeaders)
            {
                Row headerRow = new() { RowIndex = currentRowIndex };
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    headerRow.Append(CreateCell(
                        dt.Columns[i].ColumnName,
                        CellReference.CreateForOpenXml(currentRowIndex, startColIndex + (uint)i).ToString(),
                        CellValues.String));
                }
                sheetData.Append(headerRow);
                currentRowIndex++;
            }

            // Escrever dados
            foreach (DataRow dataRow in dt.Rows)
            {
                Row newRow = new Row { RowIndex = currentRowIndex };
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    object value = dataRow[i];
                    newRow.Append(CreateCell(
                        value,
                        CellReference.CreateForOpenXml(currentRowIndex, startColIndex + (uint)i).ToString(),
                        GetDataType(value)));
                }
                sheetData.Append(newRow);
                currentRowIndex++;
            }

            worksheetPart.Worksheet.Save();
        }

        private static Cell CreateCell(object value, string cellReference, CellValues dataType)
        {
            return new Cell
            {
                CellReference = cellReference,
                DataType = dataType,
                CellValue = new CellValue(ConvertValue(value))
            };
        }

        private static void WriteDataWithOpenXmlWriter(
            DataTable dt,
            WorksheetPart worksheetPart,
            uint startRowIndex,
            uint startColIndex,
            bool addHeaders)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart);
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            uint currentRowIndex = startRowIndex;

            // Escrever cabeçalhos se necessário
            if (addHeaders)
            {
                writer.WriteStartElement(new Row { RowIndex = currentRowIndex });
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    writer.WriteStartElement(new Cell
                    {
                        CellReference = CellReference.CreateForOpenXml(currentRowIndex, startColIndex + (uint)i).ToString(),
                        DataType = CellValues.String
                    });
                    writer.WriteElement(new CellValue(dt.Columns[i].ColumnName));
                    writer.WriteEndElement(); // Cell
                }
                writer.WriteEndElement(); // Row
                currentRowIndex++;
            }

            // Escrever dados
            foreach (DataRow dataRow in dt.Rows)
            {
                writer.WriteStartElement(new Row { RowIndex = currentRowIndex });
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    object value = dataRow[i];
                    writer.WriteStartElement(new Cell
                    {
                        CellReference = CellReference.CreateForOpenXml(currentRowIndex, startColIndex + (uint)i).ToString(),
                        DataType = GetDataType(value)
                    });
                    writer.WriteElement(new CellValue(ConvertValue(value)));
                    writer.WriteEndElement(); // Cell
                }
                writer.WriteEndElement(); // Row
                currentRowIndex++;
            }

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
            writer.Close();
        }

        private static CellValues GetDataType(object value)
        {
            return (value is int || value is long || value is double || value is decimal || value is float) ?
                CellValues.Number : CellValues.String;
        }


        private static string ConvertValue(object value)
        {
            if (value == null || value == DBNull.Value)
                return string.Empty;

            if (value is IFormattable formattable &&
                (value is int || value is long || value is double || value is decimal || value is float))
            {
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            }

            return value.ToString();
        }
    }
}