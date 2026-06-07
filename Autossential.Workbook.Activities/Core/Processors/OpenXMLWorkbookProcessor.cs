using Autossential.Workbook.Activities.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class OpenXMLWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        protected override CellReference ResolveCell(string address) => CellReference.OpenXml(address);

        protected override RangeReference ResolveRange(string range) => RangeReference.OpenXml(range);

        public override void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders)
        {
            WorkbookStream.Position = 0;
            using (var doc = SpreadsheetDocument.Open(WorkbookStream, true))
            {
                var wbPart = doc.WorkbookPart;
                var sheet = wbPart.GetOrCreateSheet(sheetName);
                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                var cellRef = CellReference.OpenXml(startingCell);

                int StartRow = cellRef.Row;
                int StartCol = cellRef.Col;
                uint firstRow = (uint)StartRow;
                uint lastRow = (uint)(StartRow + data.Rows.Count);

                var targetCols = Enumerable.Range(StartCol, data.Columns.Count).Select(CellReference.GetColumnName).ToHashSet();

                // Indexes all existing rows once — O(n) total
                var existingRows = sheetData.Elements<Row>().ToDictionary(r => (int)r.RowIndex.Value);

                // Removes only cells within the column range in affected rows; preserves row and cells outside the column range
                for (uint ri = firstRow; ri <= lastRow; ri++)
                {
                    if (!existingRows.TryGetValue((int)ri, out var existingRow))
                        continue;

                    existingRow.Elements<Cell>()
                               .Where(c =>
                               {
                                   var col = new string(c.CellReference?.Value?.TakeWhile(char.IsLetter).ToArray());
                                   return targetCols.Contains(col);
                               })
                               .ToList().ForEach(c => c.Remove());
                }

                // Anchors are located once to maintain SheetData order. The range cells have been removed, so the next cell (if any) is always a greater reference
                var anchor = existingRows.Values.Where(r => r.RowIndex.Value > lastRow).OrderBy(r => r.RowIndex.Value).FirstOrDefault();

                Row GetOrCreateRow(int rowIndex)
                {
                    if (existingRows.TryGetValue(rowIndex, out var existing))
                        return existing;

                    var newRow = new Row { RowIndex = (uint)rowIndex };
                    if (anchor != null)
                        sheetData.InsertBefore(newRow, anchor);
                    else
                        sheetData.AppendChild(newRow);

                    existingRows[rowIndex] = newRow;
                    return newRow;
                }

                // As the range cells have been removed, AppendChild is always safe — the only remaining cells in the row are outside the column range and have greater references, so inserting before them maintains order
                void AppendCell(Row row, Cell cell)
                {
                    var nextCell = row.Elements<Cell>().FirstOrDefault(c =>
                        string.Compare(c.CellReference?.Value, cell.CellReference?.Value,
                                       StringComparison.OrdinalIgnoreCase) > 0);
                    if (nextCell != null)
                        row.InsertBefore(cell, nextCell);
                    else
                        row.AppendChild(cell);
                }

                var (dateStyle, timeStyle, dateTimeStyle) = EnsureStyles(wbPart);

                Cell BuildCell(int rowIndex, int colIndex, object value)
                {
                    var cellRef = CellReference.GetColumnName(colIndex) + rowIndex;
                    var cell = new Cell { CellReference = cellRef };

                    switch (value)
                    {
                        case null:
                        case DBNull:
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
                            cell.StyleIndex = dt.TimeOfDay == TimeSpan.Zero
                                                ? dateStyle
                                                : dt.Date == DateTime.MinValue.Date
                                                    ? timeStyle
                                                    : dateTimeStyle;
                            break;

                        case DateTimeOffset dto:
                            cell.CellValue = new CellValue(dto.DateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                            cell.StyleIndex = dto.TimeOfDay == TimeSpan.Zero
                                                ? dateStyle
                                                : dto.Date == DateTime.MinValue.Date
                                                    ? timeStyle
                                                    : dateTimeStyle;
                            break;

                        case TimeSpan ts:
                            cell.CellValue = new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture));
                            cell.StyleIndex = timeStyle;
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
                            cell.InlineString = new InlineString(new Text(value.ToString() ?? string.Empty));
                            break;
                    }

                    return cell;
                }

                if (addHeaders)
                {
                    var headerRow = GetOrCreateRow(StartRow);
                    for (int c = 0; c < data.Columns.Count; c++)
                        AppendCell(headerRow, BuildCell(StartRow, StartCol + c, data.Columns[c].ColumnName));
                }

                for (int r = 0; r < data.Rows.Count; r++)
                {
                    int rowIndex = StartRow + 1 + r;
                    var row = GetOrCreateRow(rowIndex);
                    var dr = data.Rows[r];
                    for (int c = 0; c < data.Columns.Count; c++)
                        AppendCell(row, BuildCell(rowIndex, StartCol + c, dr[c]));
                }

                wsPart.Worksheet.Save();
            }
        }

        private static (uint DateStyle, uint TimeStyle, uint DateTimeStyle) EnsureStyles(WorkbookPart wbPart)
        {
            var stylesPart = wbPart.WorkbookStylesPart
                             ?? wbPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ?? new Stylesheet();

            stylesheet.NumberingFormats ??= new NumberingFormats();

            stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());

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

            // Built-in number format IDs:
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1

            uint dateStyle = EnsureCellFormat(14);
            uint timeStyle = EnsureCellFormat(21);
            uint dateTimeStyle = EnsureCellFormat(22);

            stylesPart.Stylesheet ??= stylesheet;
            stylesPart.Stylesheet.Save();

            return (dateStyle, timeStyle, dateTimeStyle);
        }
    }
}