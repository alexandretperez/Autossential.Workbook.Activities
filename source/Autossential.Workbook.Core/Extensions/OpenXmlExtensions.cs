using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace Autossential.Workbook.Core.Extensions
{
    internal static class OpenXmlExtensions
    {
        private static uint GenerateSheetId(this Sheets sheets)
        {
            uint maxId = 0;
            foreach (Sheet sheet in sheets.Elements<Sheet>())
            {
                if (sheet.SheetId?.Value > maxId)
                    maxId = sheet.SheetId.Value;
            }
            return maxId + 1;
        }

        public static Sheet CreateSheet(this WorkbookPart workbookPart, string sheetName)
        {
            var wsPart = workbookPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            wsPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(wsPart),
                SheetId = GenerateSheetId(sheets),
                Name = sheetName
            };
            sheets.Append(sheet);
            return sheet;
        }

        public static Sheet GetOrCreateSheet(this WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == sheetName);

            if (sheet == null)
                return CreateSheet(workbookPart, sheetName);

            return sheet;
        }

        public static bool IsNewSheet(this WorksheetPart worksheetPart)
        {
            return !worksheetPart.Worksheet.Elements<SheetData>().Any();
        }
    }
}
