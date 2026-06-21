using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Autossential.Workbook.Activities.Extensions
{
    internal static class OpenXmlExtensions
    {
        extension(WorkbookPart wbPart)
        {
            public Sheet GetOrCreateSheet(string sheetName)
            {
                var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(sheet => sheetName.Equals(sheet.Name.Value, StringComparison.OrdinalIgnoreCase));
                if (sheet == null)
                {
                    var newWorksheetPart = wbPart.AddNewPart<WorksheetPart>();
                    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                    uint sheetId = (uint)(wbPart.Workbook.Sheets!.Elements<Sheet>().Count() + 1);
                    sheet = new Sheet()
                    {
                        Id = wbPart.GetIdOfPart(newWorksheetPart),
                        SheetId = sheetId,
                        Name = sheetName
                    };
                    wbPart.Workbook.Sheets!.Append(sheet);
                }
                return sheet;
            }
        }
    }
}