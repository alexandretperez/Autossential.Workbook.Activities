using Autossential.Workbook.Activities.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Autossential.Workbook.Activities.Extensions
{
    internal static class OpenXmlExtensions
    {
        extension(WorkbookPart wbPart)
        {
            public Sheet GetOrCreateSheet(string sheetName)
            {
                var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(sheet => sheet.Name == sheetName);
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

            public (uint dateStyleIndex, uint timeStyleIndex) EnsureStyles()
            {
                var stylesPart = wbPart.WorkbookStylesPart ?? wbPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ?? new Stylesheet();
                stylesheet.NumberingFormats ??= new NumberingFormats();
                const uint dateNumFormatId = 164; // yyyy-mm-dd
                const uint timeNumFormatId = 165; // [h]:mm:ss

                void EnsureNumeringFormat(uint id, string formatCode)
                {
                    if (!stylesheet.NumberingFormats.Elements<NumberingFormat>().Any(p => p.NumberFormatId?.Value == id))
                    {
                        stylesheet.NumberingFormats.AppendChild(new NumberingFormat
                        {
                            NumberFormatId = id,
                            FormatCode = formatCode
                        });
                        stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Elements<NumberingFormat>().Count();
                    }
                }

                EnsureNumeringFormat(dateNumFormatId, "yyyy-mm-dd");
                EnsureNumeringFormat(timeNumFormatId, "[h]:mm:ss");

                stylesheet.Fonts ??= new Fonts();
                if (stylesheet.Fills == null || stylesheet.Fills.Count == 0)
                {
                    stylesheet.Fills = new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                    );
                }

                stylesheet.Borders ??= new Borders(new Border());
                stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
                stylesheet.CellFormats ??= new CellFormats(new CellFormat());

                uint EnsureCellFormat(uint numFmtId)
                {
                    var existing = stylesheet.CellFormats.Elements<CellFormat>()
                                             .Select((cf, i) => (cf, i))
                                             .FirstOrDefault(x => x.cf.NumberFormatId?.Value == numFmtId);
                    if (existing.cf != null)
                        return (uint)existing.i;

                    stylesheet.CellFormats.AppendChild(new CellFormat
                    {
                        NumberFormatId = numFmtId,
                        ApplyNumberFormat = true
                    });
                    stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats
                                                                   .Elements<CellFormat>().Count();
                    return stylesheet.CellFormats.Count - 1;
                }

                uint dateStyleIndex = EnsureCellFormat(dateNumFormatId);
                uint timeStyleIndex = EnsureCellFormat(timeNumFormatId);

                stylesPart.Stylesheet ??= stylesheet;
                stylesPart.Stylesheet.Save();

                return (dateStyleIndex, timeStyleIndex);
            }
        }


        extension(SheetData sheetData)
        {
            public Row GetOrCreateRow(uint rowIndex, Row anchorRow, Dictionary<uint, Row> existingRows)
            {
                if (existingRows.TryGetValue(rowIndex, out var existingRow))
                    return existingRow;

                var newRow = new Row { RowIndex = rowIndex };
                if (anchorRow == null)
                {
                    sheetData.AppendChild(newRow);
                }
                else
                {
                    sheetData.InsertBefore(newRow, anchorRow);
                }

                existingRows[rowIndex] = newRow;
                return newRow;
            }
        }

        extension(Row row)
        {
            public void BuildAndAppendCell(int rowIndex, int colIndex, object cellValue, uint dateStyleIndex, uint timeStyleIndex)
            {
                var cellRef = CellReference.GetColumnName(colIndex) + rowIndex;
                var cell = new Cell { CellReference = cellRef };

                switch (cellValue)
                {
                    case null or DBNull:
                        break;
                    case string s:
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(s));
                        break;

                    case bool b:
                        cell.DataType = CellValues.Boolean;
                        cell.CellValue = new CellValue(b ? "1" : "0");
                        break;

                    case DateTime dt:
                        cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = dateStyleIndex;
                        break;

                    case DateTimeOffset dto:
                        cell.CellValue = new CellValue(dto.DateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = dateStyleIndex;
                        break;

                    case TimeSpan ts:
                        cell.CellValue = new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture));
                        cell.StyleIndex = timeStyleIndex;
                        break;

                    case Guid g:
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(g.ToString()));
                        break;

                    case double d:
                        cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
                        break;
                    case float f:
                        cell.CellValue = new CellValue(((double)f).ToString(CultureInfo.InvariantCulture));
                        break;
                    case decimal dec:
                        cell.CellValue = new CellValue(dec.ToString(CultureInfo.InvariantCulture));
                        break;
                    case int i:
                        cell.CellValue = new CellValue(i.ToString(CultureInfo.InvariantCulture));
                        break;
                    case long l:
                        cell.CellValue = new CellValue(l.ToString(CultureInfo.InvariantCulture));
                        break;
                    case short sh:
                        cell.CellValue = new CellValue(sh.ToString(CultureInfo.InvariantCulture));
                        break;
                    case byte by:
                        cell.CellValue = new CellValue(by.ToString(CultureInfo.InvariantCulture));
                        break;

                    default:
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(cellValue.ToString() ?? string.Empty));
                        break;
                }

                var nextCell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference?.Value, cell.CellReference?.Value, StringComparison.OrdinalIgnoreCase) > 0);
                if (nextCell == null)
                    row.AppendChild(cell);
                else
                    row.InsertBefore(cell, nextCell);
            }
        }
    }
}