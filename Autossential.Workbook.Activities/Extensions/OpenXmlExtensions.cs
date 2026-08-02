using Autossential.Workbook.Activities.Core;
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

        extension(SheetData sheetData)
        {
            public List<KeyValuePair<int, Row>>.Enumerator BuildRowEnumerator()
            {
                int previousRowIndex = 0;
                var list = new List<KeyValuePair<int, Row>>();
                foreach (var row in sheetData.Elements<Row>())
                {
                    int rowIndex = row.RowIndex?.Value is uint explicitRow
                      ? (int)explicitRow
                      : previousRowIndex + 1;

                    list.Add(new KeyValuePair<int, Row>(rowIndex, row));
                    previousRowIndex = rowIndex;
                }
                list.Sort((a, b) => a.Key.CompareTo(b.Key));
                return list.GetEnumerator();
            }
        }


        extension(Row row)
        {
            public List<KeyValuePair<int, Cell>>.Enumerator BuildCellEnumerator()
            {
                int previousColIndex = 0;
                var list = new List<KeyValuePair<int, Cell>>();
                foreach (var cell in row.Elements<Cell>())
                {
                    int colIndex = cell.CellReference?.Value is string cellRef
                        ? new OpenXmlCellReference(cellRef).Col
                        : previousColIndex + 1;

                    list.Add(new KeyValuePair<int, Cell>(colIndex, cell));
                    previousColIndex = colIndex;
                }
                list.Sort((a, b) => a.Key.CompareTo(b.Key));
                return list.GetEnumerator();
            }
        }
    }
}