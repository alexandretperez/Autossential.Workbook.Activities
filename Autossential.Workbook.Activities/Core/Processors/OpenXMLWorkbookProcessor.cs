using Autossential.Workbook.Activities.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class OpenXMLWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        private SpreadsheetDocument GetWorkbook()
        {
            WorkbookStream.Position = 0;
            return SpreadsheetDocument.Open(WorkbookStream, true);
        }

        protected override CellReference ResolveCell(string address) => new OpenXmlCellReference(address);

        protected override RangeReference ResolveRange(string range) => new OpenXmlRangeReference(range);

        public override void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders)
        {
            ValidateSheetName(sheetName);
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheet = wbPart.GetOrCreateSheet(sheetName);
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            var cellRef = ResolveCell(startingCell);

            int startRow = cellRef.Row;
            int startCol = cellRef.Col;
            uint firstRow = (uint)startRow;
            uint lastRow = (uint)(startRow + data.Rows.Count);

            var targetCols = Enumerable.Range(startCol, data.Columns.Count)
                                       .Select(CellReference.GetColumnName)
                                       .ToHashSet();

            var existingRows = sheetData.Elements<Row>()
                                        .ToDictionary(r => (int)r.RowIndex.Value);

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
                           .ToList()
                           .ForEach(c => c.Remove());
            }

            var anchor = existingRows.Values
                                     .Where(r => r.RowIndex.Value > lastRow)
                                     .OrderBy(r => r.RowIndex.Value)
                                     .FirstOrDefault();

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

            // Loads SharedStrings only once and builds the index in memory to avoid repeated linear searches while writing the range
            var sst = GetOrCreateSharedStringTable(wbPart);
            var sstIndex = BuildSharedStringIndex(sst);

            var (dateStyle, timeStyle, dateTimeStyle) = EnsureStyles(wbPart);

            Cell BuildCell(int rowIndex, int colIndex, object value)
            {
                var cellAddress = CellReference.GetColumnName(colIndex) + rowIndex;
                var cell = new Cell { CellReference = cellAddress };
                UpdateCell(cell, value, dateStyle, timeStyle, dateTimeStyle, sst, sstIndex);
                return cell;
            }

            if (addHeaders)
            {
                var headerRow = GetOrCreateRow(startRow);
                for (int c = 0; c < data.Columns.Count; c++)
                    AppendCell(headerRow, BuildCell(startRow, startCol + c, data.Columns[c].ColumnName));
                startRow++;
            }

            for (int r = 0; r < data.Rows.Count; r++)
            {
                int rowIndex = startRow + r;
                var row = GetOrCreateRow(rowIndex);
                var dr = data.Rows[r];
                for (int c = 0; c < data.Columns.Count; c++)
                    AppendCell(row, BuildCell(rowIndex, startCol + c, dr[c]));
            }

            // Updates the SST count before save it
            sst.Count = (uint)sstIndex.Count;
            sst.UniqueCount = (uint)sstIndex.Count;
            wbPart.SharedStringTablePart.SharedStringTable.Save();
            wsPart.Worksheet.Save();

            RemoveDefaultSheetIfNeed(wbPart, sheetName);
        }

        public override void WriteCell(string sheetName, string address, object value)
        {
            ValidateSheetName(sheetName);
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheet = wbPart.GetOrCreateSheet(sheetName);

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            var cellRef = ResolveCell(address);
            var rowIndex = cellRef.Row;

            var existingRow = sheetData.Elements<Row>()
                                       .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
            if (existingRow == null)
            {
                existingRow = new Row { RowIndex = (uint)rowIndex };
                var anchor = sheetData.Elements<Row>()
                                      .FirstOrDefault(r => r.RowIndex?.Value > (uint)rowIndex);
                if (anchor != null)
                    sheetData.InsertBefore(existingRow, anchor);
                else
                    sheetData.AppendChild(existingRow);
            }

            var cell = existingRow.Elements<Cell>()
                                  .FirstOrDefault(c => c.CellReference?.Value == address);
            if (cell == null)
            {
                cell = new Cell { CellReference = address };
                var nextCell = existingRow.Elements<Cell>().FirstOrDefault(c =>
                    string.Compare(c.CellReference?.Value, address,
                                   StringComparison.OrdinalIgnoreCase) > 0);
                if (nextCell != null)
                    existingRow.InsertBefore(cell, nextCell);
                else
                    existingRow.AppendChild(cell);
            }

            cell.RemoveAllChildren();
            cell.DataType = null;
            cell.StyleIndex = null;

            var sst = GetOrCreateSharedStringTable(wbPart);
            var sstIndex = BuildSharedStringIndex(sst);
            var (dateStyle, timeStyle, dateTimeStyle) = EnsureStyles(wbPart);

            UpdateCell(cell, value, dateStyle, timeStyle, dateTimeStyle, sst, sstIndex);

            sst.Count = (uint)sstIndex.Count;
            sst.UniqueCount = (uint)sstIndex.Count;
            wbPart.SharedStringTablePart.SharedStringTable.Save();
            wsPart.Worksheet.Save();

            RemoveDefaultSheetIfNeed(wbPart, sheetName);
        }

        private static SharedStringTable GetOrCreateSharedStringTable(WorkbookPart wbPart)
        {
            var sstPart = wbPart.SharedStringTablePart
                          ?? wbPart.AddNewPart<SharedStringTablePart>();

            sstPart.SharedStringTable ??= new SharedStringTable();
            return sstPart.SharedStringTable;
        }

        private static Dictionary<string, int> BuildSharedStringIndex(SharedStringTable sst)
        {
            var index = new Dictionary<string, int>(StringComparer.Ordinal);
            int i = 0;
            foreach (var item in sst.Elements<SharedStringItem>())
            {
                var text = item.InnerText;
                index.TryAdd(text, i);
                i++;
            }
            return index;
        }

        private static int GetOrAddSharedString(SharedStringTable sst,
                                                Dictionary<string, int> sstIndex,
                                                string value)
        {
            if (sstIndex.TryGetValue(value, out var idx))
                return idx;

            sst.AppendChild(new SharedStringItem(new Text(value)));
            idx = sstIndex.Count;
            sstIndex[value] = idx;
            return idx;
        }

        private static void UpdateCell(Cell cell, object value,
                                       uint dateStyle, uint timeStyle, uint dateTimeStyle,
                                       SharedStringTable sst, Dictionary<string, int> sstIndex)
        {
            switch (value)
            {
                case null:
                case DBNull:
                    break;

                case string s:
                    if (s.StartsWith('='))
                    {
                        cell.CellFormula = new CellFormula(s[1..]);
                        cell.CellValue = new CellValue();
                    }
                    else
                    {
                        cell.DataType = CellValues.SharedString;
                        cell.CellValue = new CellValue(GetOrAddSharedString(sst, sstIndex, s).ToString());
                    }
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
                                    : dto.Date == DateTimeOffset.MinValue.Date
                                    ? timeStyle
                                    : dateTimeStyle;
                    break;

                case TimeSpan ts:
                    cell.CellValue = new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = timeStyle;
                    break;

                case Guid g:
                    cell.DataType = CellValues.SharedString;
                    cell.CellValue = new CellValue(GetOrAddSharedString(sst, sstIndex, g.ToString()).ToString());
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
                    cell.DataType = CellValues.SharedString;
                    cell.CellValue = new CellValue(GetOrAddSharedString(sst, sstIndex, value.ToString() ?? string.Empty).ToString());
                    break;
            }
        }

        private static (uint DateStyle, uint TimeStyle, uint DateTimeStyle) EnsureStyles(WorkbookPart wbPart)
        {
            var stylesPart = wbPart.WorkbookStylesPart
                             ?? wbPart.AddNewPart<WorkbookStylesPart>();

            var stylesheet = stylesPart.Stylesheet ?? new Stylesheet();

            stylesheet.NumberingFormats ??= new NumberingFormats();

            stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font(
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" }
            ));

            if (stylesheet.Fills == null || !stylesheet.Fills.Elements<Fill>().Any())
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

            uint dateStyle = EnsureCellFormat(14);
            uint timeStyle = EnsureCellFormat(21);
            uint dateTimeStyle = EnsureCellFormat(22);

            stylesPart.Stylesheet ??= stylesheet;
            stylesPart.Stylesheet.Save();

            return (dateStyle, timeStyle, dateTimeStyle);
        }

        protected override void CreateNew()
        {
            using var doc = SpreadsheetDocument.Create(WorkbookStream, SpreadsheetDocumentType.Workbook, autoSave: false);

            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            var sstPart = wbPart.AddNewPart<SharedStringTablePart>();
            sstPart.SharedStringTable = new SharedStringTable
            {
                Count = 0,
                UniqueCount = 0
            };

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = SetDefaultSheetName()
            });

            doc.Save();
        }

        private void RemoveDefaultSheetIfNeed(WorkbookPart wbPart, string sheetName)
        {
            var defaultSheetName = ConsumeDefaultSheetName();
            if (defaultSheetName == null || defaultSheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return;

            var sheet = wbPart.Workbook.Sheets
                              .Elements<Sheet>()
                              .FirstOrDefault(s => s.Name?.Value == defaultSheetName); // case-insensitive is being handled above

            if (sheet == null)
                return;

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);
            wbPart.DeletePart(wsPart);
            sheet.Remove();

            wbPart.Workbook.Save();
        }

        public override void DeleteSheet(string sheetName)
        {
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;

            var sheets = wbPart.Workbook.Sheets.Elements<Sheet>();
            var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
                return;

            if (sheets.Count() == 1)
                throw new InvalidOperationException($"Cannot delete the only sheet \"{sheetName}\" in the workbook.");

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);
            wbPart.DeletePart(wsPart);
            sheet.Remove();

            wbPart.Workbook.Save();
        }

        public override void InsertSheet(string sheetName, int? position = null)
        {
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;

            var existingSheet = wbPart.Workbook.Sheets
                .Elements<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

            if (existingSheet != null)
                throw new InvalidOperationException($"A sheet with name '{sheetName}' already exists.");

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            uint sheetId = 1;

            var sheetElements = wbPart.Workbook.Sheets.Elements<Sheet>();
            if (sheetElements.Any())
                sheetId = sheetElements.Max(s => s.SheetId.Value) + 1;

            var newSheet = new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = sheetId,
                Name = sheetName
            };

            var sheets = wbPart.Workbook.Sheets;
            if (position.HasValue && position.Value > 0 && position.Value <= sheets.Count())
            {
                var refSheet = sheets.Elements<Sheet>().ElementAt(position.Value - 1);
                refSheet.InsertBeforeSelf(newSheet);
            }
            else
            {
                sheets.Append(newSheet);
            }

            wbPart.Workbook.Save();
        }

        public override void RenameSheet(string fromSheetName, string toSheetName)
        {
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Sheets.Elements<Sheet>();
            var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, fromSheetName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException($"No sheet with name '{fromSheetName}' was found.");

            if (sheet.Name.Value == toSheetName)
                return;

            var anotherSheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, toSheetName, StringComparison.OrdinalIgnoreCase));
            if (anotherSheet is not null && anotherSheet.Id.Value != sheet.Id.Value)
                throw new InvalidOperationException($"Another sheet with name '{toSheetName}' already exists in the workbook.");

            sheet.Name = toSheetName;
            wbPart.Workbook.Save();
        }

        public override void FreezePanes(string sheetName, int colsToFreeze, int rowsToFreeze)
        {
            ValidateSheetName(sheetName);

            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;

            var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
                return;

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
            var worksheet = wsPart.Worksheet;
            var sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews is null)
            {
                sheetViews = new SheetViews();
                worksheet.InsertAt(sheetViews, 0);
            }

            var sheetView = sheetViews.GetFirstChild<SheetView>();
            if (sheetView is null)
            {
                sheetView = new SheetView { WorkbookViewId = 0 };
                sheetViews.Append(sheetView);
            }

            sheetView.RemoveAllChildren<Pane>();
            sheetView.RemoveAllChildren<Selection>();

            var freezeCols = colsToFreeze > 0;
            var freezeRows = rowsToFreeze > 0;

            if (!freezeCols && !freezeRows)
            {
                worksheet.Save();
                return;
            }

            var topLeftCell = new OpenXmlCellReference(colsToFreeze + 1, rowsToFreeze + 1).ToAddress();

            var activePane = (freezeCols, freezeRows) switch
            {
                (true, true) => PaneValues.BottomRight,
                (false, true) => PaneValues.BottomLeft,
                (true, false) => PaneValues.TopRight,
                _ => PaneValues.TopLeft
            };

            var pane = new Pane
            {
                HorizontalSplit = colsToFreeze,
                VerticalSplit = rowsToFreeze,
                TopLeftCell = topLeftCell,
                ActivePane = activePane,
                State = PaneStateValues.Frozen
            };

            sheetView.Append(pane);
            sheetView.Append(new Selection
            {
                Pane = activePane,
                ActiveCell = topLeftCell,
                SequenceOfReferences = new ListValue<StringValue>
                {
                    InnerText = topLeftCell
                }
            });

            worksheet.Save();
        }

        public override void HideSheet(string sheetName) => 
            ToggleSheetState(sheetName, SheetStateValues.Hidden);

        public override void UnhideSheet(string sheetName) => 
            ToggleSheetState(sheetName, SheetStateValues.Visible);

        private void ToggleSheetState(string sheetName, SheetStateValues state)
        {
            ValidateSheetName(sheetName);
            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Sheets;
            foreach (Sheet sheet in sheets.Cast<Sheet>())
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    sheet.State = state;
            }
            wbPart.Workbook.Save();
        }
    }
}