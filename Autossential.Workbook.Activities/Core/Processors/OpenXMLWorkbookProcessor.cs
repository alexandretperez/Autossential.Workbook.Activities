using Autossential.Workbook.Activities.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MathNet.Numerics.Distributions;
using System.Data;
using System.Globalization;
using System.Runtime.Intrinsics.Arm;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class OpenXMLWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        public override bool IsOpenXML => true;

        public override bool IsBIFF8 => false;

        private SpreadsheetDocument GetWorkbook()
        {
            WorkbookStream.Position = 0;
            return SpreadsheetDocument.Open(WorkbookStream, true);
        }

        protected override CellRef ResolveCell(string address) => CellRef.Parse(address.AsSpan());

        protected override RangeRef ResolveRange(string address) => RangeRef.Parse(address.AsSpan());

        public override void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders)
        {
            ValidateSheetName(sheetName);

            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheet = wbPart.GetOrCreateSheet(sheetName);
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.RemoveDefaultEmptyRows();

            var cellRef = ResolveCell(startingCell);

            int startRow = cellRef.Row;
            int startCol = cellRef.Col;

            // Loads SharedStrings only once and builds the index in memory to avoid repeated linear searches while writing the range
            var sst = GetOrCreateSharedStringTable(wbPart);
            var sstIndex = BuildSharedStringIndex(sst);

            var (dateStyle, timeStyle, dateTimeStyle) = EnsureStyles(wbPart);

            var rows = sheetData.BuildRowEnumerator();
            KeyValuePair<int, Row>? currentRow = rows.MoveNext() ? rows.Current : null;

            // Pre-computed column names for the range being written, avoiding repeated
            // CellReference.GetColumnName + string concatenation per cell
            var columnNames = new string[data.Columns.Count];
            for (int i = 0; i < columnNames.Length; i++)
                columnNames[i] = CellRef.GetColumnName(startCol + i);

            static string BuildCellReference(string columnName, int rowIndex)
            {
                var rowDigits = (int)Math.Floor(Math.Log10(rowIndex)) + 1;
                return string.Create(columnName.Length + rowDigits, (columnName, rowIndex), (span, state) =>
                {
                    state.columnName.AsSpan().CopyTo(span);
                    state.rowIndex.TryFormat(span[state.columnName.Length..], out _);
                });
            }

            Row GetOrCreateRow(int rowIndex)
            {
                while (currentRow.HasValue && currentRow.Value.Key < rowIndex)
                    currentRow = rows.MoveNext() ? rows.Current : null;

                if (currentRow.HasValue && currentRow.Value.Key == rowIndex)
                    return currentRow.Value.Value;

                var newRow = new Row { RowIndex = (uint)rowIndex };
                if (currentRow == null)
                    sheetData.AppendChild(newRow);
                else
                    sheetData.InsertBefore(newRow, currentRow.Value.Value);

                return newRow;
            }

            void UpdateOrCreateCell(Row row, List<KeyValuePair<int, Cell>>.Enumerator remaining, ref KeyValuePair<int, Cell>? current, int colIndex, int rowIndex, object value)
            {
                while (current.HasValue && current.Value.Key < colIndex)
                    current = remaining.MoveNext() ? remaining.Current : null;

                Cell cell;
                if (current.HasValue && current.Value.Key == colIndex)
                {
                    cell = current.Value.Value;
                }
                else
                {
                    var cellReference = BuildCellReference(columnNames[colIndex - startCol], rowIndex);
                    cell = new Cell { CellReference = cellReference };
                    if (current == null)
                        row.AppendChild(cell);
                    else
                        row.InsertBefore(cell, current.Value.Value);
                }

                UpdateCell(cell, value, dateStyle, timeStyle, dateTimeStyle, sst, sstIndex);
            }

            if (addHeaders)
            {
                var headerRow = GetOrCreateRow(startRow);
                var cells = headerRow.BuildCellEnumerator();
                KeyValuePair<int, Cell>? current = cells.MoveNext() ? cells.Current : null;
                for (int i = 0; i < data.Columns.Count; i++)
                    UpdateOrCreateCell(headerRow, cells, ref current, startCol + i, startRow, data.Columns[i].ColumnName);

                startRow++;
            }

            for (int i = 0; i < data.Rows.Count; i++)
            {
                var rowIndex = startRow + i;
                var row = GetOrCreateRow(rowIndex);
                var cells = row.BuildCellEnumerator();
                KeyValuePair<int, Cell>? current = cells.MoveNext() ? cells.Current : null;

                var dr = data.Rows[i];
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    var colIndex = startCol + j;
                    UpdateOrCreateCell(row, cells, ref current, colIndex, rowIndex, dr[j]);
                }
            }

            sst.UniqueCount = (uint)sstIndex.Count;
            wbPart.SharedStringTablePart.SharedStringTable.Save();
            wsPart.Worksheet.SetActiveCellToA1();
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

            var rows = sheetData.BuildRowEnumerator();
            KeyValuePair<int, Row>? currentRow = rows.MoveNext() ? rows.Current : null;
            while (currentRow.HasValue && currentRow.Value.Key < rowIndex)
                currentRow = rows.MoveNext() ? rows.Current : null;

            Row row;
            if (currentRow.HasValue && currentRow.Value.Key == rowIndex)
            {
                row = currentRow.Value.Value;
            }
            else
            {
                row = new Row { RowIndex = (uint)rowIndex };
                if (currentRow == null)
                    sheetData.AppendChild(row);
                else
                    sheetData.InsertBefore(row, currentRow.Value.Value);
            }

            var cells = row.BuildCellEnumerator();
            KeyValuePair<int, Cell>? currentCell = cells.MoveNext() ? cells.Current : null;

            while (currentCell.HasValue && currentCell.Value.Key < cellRef.Col)
                currentCell = cells.MoveNext() ? cells.Current : null;

            Cell cell;
            if (currentCell.HasValue && currentCell.Value.Key == cellRef.Col)
            {
                cell = currentCell.Value.Value;
            }
            else
            {
                cell = new Cell { CellReference = address };
                if (currentCell == null)
                    row.AppendChild(cell);
                else
                    row.InsertBefore(cell, currentCell.Value.Value);
            }

            var sst = GetOrCreateSharedStringTable(wbPart);
            var sstIndex = BuildSharedStringIndex(sst);
            var (dateStyle, timeStyle, dateTimeStyle) = EnsureStyles(wbPart);

            UpdateCell(cell, value, dateStyle, timeStyle, dateTimeStyle, sst, sstIndex);

            sst.UniqueCount = (uint)sstIndex.Count;
            wbPart.SharedStringTablePart.SharedStringTable.Save();
            wsPart.Worksheet.SetActiveCellToA1();
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
            cell.RemoveAllChildren();
            cell.DataType = null;

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

            var topLeftCell = new CellRef(colsToFreeze + 1, rowsToFreeze + 1).GetAddress();

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

        public override void DeleteColumns(string sheetName, string references)
        {
            ValidateSheetName(sheetName);
            var positions = ResolveColumnsReferences(references);

            if (positions.Count == 0)
                return;

            var columnsDesc = new List<int>(positions);
            columnsDesc.Sort((a, b) => b.CompareTo(a));

            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Sheets.Elements<Sheet>();
            var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                     ?? throw new InvalidOperationException($"No sheet with name '{sheetName}' was found.");

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null)
                return;

            var columnsAsc = new List<int>(columnsDesc);
            columnsAsc.Reverse();

            var rows = sheetData.BuildRowEnumerator();
            while (rows.MoveNext())
            {
                var (ri, row) = rows.Current;
                var removed = false;

                var cells = row.BuildCellEnumerator();
                while (cells.MoveNext())
                {
                    var (ci, cell) = cells.Current;
                    if (!columnsDesc.Contains(ci))
                        continue;

                    cell.Remove();
                    removed = true;
                }

                if (!removed)
                    continue;

                int shift = 0;
                int dp = 0;
                cells = row.BuildCellEnumerator();
                while (cells.MoveNext())
                {
                    var (ci, cell) = cells.Current;
                    while (dp < columnsAsc.Count && columnsAsc[dp] < ci)
                    {
                        dp++;
                        shift++;
                    }
                    if (shift == 0)
                        continue;

                    var newIndex = ci - shift;
                    if (cell.CellReference?.HasValue == true)
                        cell.CellReference = new CellRef(newIndex, ri).GetAddress();
                }
            }

            var worksheet = wsPart.Worksheet;

            AdjustMergedCells(worksheet, columnsDesc, false);
            UpdateSheetDimension(worksheet, sheetData);

            worksheet.SetActiveCellToA1();
            worksheet.Save();
        }

        private static void AdjustMergedCells(Worksheet worksheet, List<int> deletedPositions, bool isRowAxis)
        {
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault()?.Elements<MergeCell>().ToList() ?? [];

            if (mergeCells.Count == 0)
                return;

            int rowShift = isRowAxis ? 1 : 0;
            int colShift = isRowAxis ? 0 : 1;

            foreach (var pos in deletedPositions)
            {
                int i = 0;
                while (i < mergeCells.Count)
                {
                    var mergeCell = mergeCells[i];
                    var range = RangeRef.Parse(mergeCell.Reference);
                    var start = range.Start;
                    var end = range.End;

                    if (isRowAxis ? pos < start.Row : pos < start.Col)
                    {
                        start = new CellRef(range.Start.Col - colShift, range.Start.Row - rowShift);
                        end = new CellRef(range.End.Col - colShift, range.End.Row - rowShift);
                    }
                    else if (isRowAxis ? pos <= end.Row : pos <= end.Col)
                    {
                        end = new CellRef(range.End.Col - colShift, range.End.Row - rowShift);
                    }

                    if (start == end)
                    {
                        mergeCells.Remove(mergeCell);
                        mergeCell.Remove();
                        continue;
                    }
                    else
                    {
                        mergeCell.Reference = new RangeRef(start, end).GetAddress();
                    }

                    i++;
                }
            }
        }

        public override void DeleteRows(string sheetName, string references)
        {
            ValidateSheetName(sheetName);

            var positions = ResolveRowsReferences(references);
            if (positions.Count == 0)
                return;

            var rowsDesc = new List<int>(positions);
            rowsDesc.Sort((a, b) => b.CompareTo(a));

            using var doc = GetWorkbook();
            var wbPart = doc.WorkbookPart;
            var sheets = wbPart.Workbook.Sheets.Elements<Sheet>();
            var sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                     ?? throw new InvalidOperationException($"No sheet with name '{sheetName}' was found.");

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null)
                return;

            var removed = false;
            var allRows = sheetData.BuildRowEnumerator();

            while (allRows.MoveNext())
            {
                var (ri, row) = allRows.Current;
                if (rowsDesc.Contains(ri))
                {
                    removed = true;
                    row.Remove();
                    continue;
                }
            }

            if (!removed)
                return;

            var rowsAsc = new List<int>(rowsDesc);
            rowsAsc.Reverse();

            int shift = 0;
            int dp = 0;

            var deletedRow = 0;

            allRows = sheetData.BuildRowEnumerator();
            while (allRows.MoveNext())
            {
                var (ri, row) = allRows.Current;

                while (dp < rowsAsc.Count && rowsAsc[dp] < ri)
                {
                    deletedRow = rowsAsc[dp];
                    dp++;
                    shift++;
                }

                if (shift == 0)
                    continue;

                var newIndex = ri - shift;

                if (row.RowIndex?.HasValue == true)
                    row.RowIndex = (uint)newIndex;

                var cells = row.BuildCellEnumerator();
                while (cells.MoveNext())
                {
                    var (ci, cell) = cells.Current;
                    if (cell.CellReference?.HasValue == true)
                        cell.CellReference = new CellRef(ci, newIndex).GetAddress();
                }
            }

            var worksheet = wsPart.Worksheet;

            AdjustMergedCells(worksheet, rowsDesc, true);
            UpdateSheetDimension(worksheet, sheetData);

            worksheet.SetActiveCellToA1();
            worksheet.Save();
        }

        private static void UpdateSheetDimension(Worksheet worksheet, SheetData sheetData)
        {
            var rows = sheetData.BuildRowEnumerator();

            int minRow = int.MaxValue, maxRow = int.MinValue;
            int minCol = int.MaxValue, maxCol = int.MinValue;
            var hasData = false;

            while (rows.MoveNext())
            {
                var (ri, row) = rows.Current;
                var cells = row.BuildCellEnumerator();
                var rowHasCells = false;

                while (cells.MoveNext())
                {
                    var (ci, _) = cells.Current;
                    hasData = true;
                    rowHasCells = true;
                    if (ci < minCol) minCol = ci;
                    if (ci > maxCol) maxCol = ci;
                }

                if (rowHasCells)
                {
                    if (ri < minRow) minRow = ri;
                    if (ri > maxRow) maxRow = ri;
                }
            }

            worksheet.SheetDimension ??= new SheetDimension();

            worksheet.SheetDimension.Reference = hasData
                ? $"{new CellRef(minCol, minRow).GetAddress()}:{new CellRef(maxCol, maxRow).GetAddress()}"
                : "A1";
        }
    }
}