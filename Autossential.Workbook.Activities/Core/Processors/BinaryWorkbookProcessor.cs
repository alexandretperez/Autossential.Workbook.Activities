using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Activities.Presentation.ViewState;
using System.Data;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class BinaryWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        private HSSFWorkbook GetWorkbook()
        {
            var editStream = new MemoryStream();
            WorkbookStream.Position = 0;
            WorkbookStream.CopyTo(editStream);
            editStream.Position = 0;

            return new HSSFWorkbook(editStream);
        }

        private void FlushWorkbook(HSSFWorkbook workbook)
        {
            WorkbookStream.Position = 0;
            WorkbookStream.SetLength(0);
            workbook.Write(WorkbookStream, true);
        }

        protected override CellReference ResolveCell(string address) => new BIFF8CellReference(address);

        protected override RangeReference ResolveRange(string range) => new BIFF8RangeReference(range);

        public override void WriteCell(string sheetName, string address, object value)
        {
            ValidateSheetName(sheetName);
            using var wb = GetWorkbook();
            var sh = wb.GetSheet(sheetName) ?? wb.CreateSheet(sheetName);

            var (dateStyle, timeStyle, dateTimeStyle) = GetCellStyles(wb);

            var cellRef = ResolveCell(address);
            var colLetter = CellReference.GetColumnName(cellRef.Col);
            var rowIndex = cellRef.Row;

            var rowIdx = rowIndex - 1; // 0-based
            var colIdx = cellRef.Col - 1; // 0-based

            var row = sh.GetRow(rowIdx) ?? sh.CreateRow(rowIdx);
            var cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);

            SetCell(cell, value, dateStyle, timeStyle, dateTimeStyle);

            RemoveDefaultSheetIfNeed(wb, sheetName);
            FlushWorkbook(wb);
        }

        public override void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders)
        {
            ValidateSheetName(sheetName);
            using var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName) ?? wb.CreateSheet(sheetName);

            var (dateStyle, timeStyle, dateTimeStyle) = GetCellStyles(wb);

            var cellRef = ResolveCell(startingCell);
            var colLetter = CellReference.GetColumnName(cellRef.Col);
            var rowIndex = cellRef.Row;

            var startRow = rowIndex - 1; // 0-based
            var startCol = cellRef.Col - 1; // 0-based

            IRow GetOrCreateRow(int index) => sheet.GetRow(index) ?? sheet.CreateRow(index);

            if (addHeaders)
            {
                var headerRow = GetOrCreateRow(startRow);
                for (int c = 0; c < data.Columns.Count; c++)
                    SetCell(headerRow.CreateCell(startCol + c), data.Columns[c].ColumnName, dateStyle, timeStyle, dateTimeStyle);

                startRow++;
            }

            for (int r = 0; r < data.Rows.Count; r++)
            {
                var row = GetOrCreateRow(startRow + r);
                var dr = data.Rows[r];
                for (int c = 0; c < data.Columns.Count; c++)
                    SetCell(row.CreateCell(startCol + c), dr[c], dateStyle, timeStyle, dateTimeStyle);
            }

            RemoveDefaultSheetIfNeed(wb, sheetName);
            FlushWorkbook(wb);
        }

        private static void SetCell(ICell cell, object value, ICellStyle dateStyle, ICellStyle timeStyle, ICellStyle dateTimeStyle)
        {
            cell.CellStyle = null; // reset

            switch (value)
            {
                case null:
                case DBNull:
                    cell.SetBlank();
                    break;

                case string s:
                    if (s.StartsWith('='))
                    {
                        cell.SetCellFormula(s[1..]);
                    }
                    {
                        cell.SetCellValue(s);
                    }
                    break;

                case bool b:
                    cell.SetCellValue(b);
                    break;

                case DateTime dt:
                    cell.SetCellValue(dt);
                    cell.CellStyle = dt.TimeOfDay == TimeSpan.Zero
                                   ? dateStyle
                                   : dt.Date == DateTime.MinValue.Date
                                   ? timeStyle
                                   : dateTimeStyle;
                    break;

                case DateTimeOffset dto:
                    cell.SetCellValue(dto.DateTime);
                    cell.CellStyle = dto.TimeOfDay == TimeSpan.Zero
                                   ? dateStyle
                                   : dto.Date == DateTimeOffset.MinValue.Date
                                   ? timeStyle
                                   : dateTimeStyle;
                    break;

                case TimeSpan ts:
                    cell.SetCellValue(ts.TotalDays);
                    cell.CellStyle = timeStyle;
                    break;

                case Guid g:
                    cell.SetCellValue(g.ToString());
                    break;

                case double d:
                    cell.SetCellValue(d);
                    break;

                case float f:
                    cell.SetCellValue((double)f);
                    break;

                case decimal dec:
                    cell.SetCellValue((double)dec);
                    break;

                case int i:
                    cell.SetCellValue(i);
                    break;

                case long l:
                    cell.SetCellValue(l);
                    break;

                case short sh:
                    cell.SetCellValue(sh);
                    break;

                case byte by:
                    cell.SetCellValue(by);
                    break;

                default:
                    cell.SetCellValue(value.ToString() ?? string.Empty);
                    break;
            }
        }

        private static (ICellStyle DateStyle, ICellStyle TimeStyle, ICellStyle DateTimeStyle) GetCellStyles(IWorkbook wb)
        {
            var dateStyle = wb.CreateCellStyle();
            dateStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat(BuiltinFormats.ShortDate);

            var timeStyle = wb.CreateCellStyle();
            timeStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat(BuiltinFormats.Time24);

            var dateTimeStyle = wb.CreateCellStyle();
            dateTimeStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat(BuiltinFormats.DateTime);

            return (dateStyle, timeStyle, dateTimeStyle);
        }

        protected override void CreateNew()
        {
            using var wb = new HSSFWorkbook();
            wb.CreateSheet(SetDefaultSheetName());
            wb.Write(WorkbookStream, leaveOpen: true);
        }

        private void RemoveDefaultSheetIfNeed(IWorkbook wb, string sheetName)
        {
            var defaultSheetName = ConsumeDefaultSheetName();
            if (defaultSheetName == null || defaultSheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return;

            var index = wb.GetSheetIndex(defaultSheetName);
            if (index < 0)
                return;

            wb.RemoveSheetAt(index);
        }

        public override void DeleteSheet(string sheetName)
        {
            using var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet == null)
                return;

            if (wb.NumberOfSheets == 1)
                throw new InvalidOperationException($"Cannot delete the only sheet \"{sheetName}\" in the workbook.");

            var sheetIndex = wb.GetSheetIndex(sheet);
            wb.RemoveSheetAt(sheetIndex);

            FlushWorkbook(wb);
        }

        public override void InsertSheet(string sheetName, int? position = null)
        {
            using var wb = GetWorkbook();
            if (wb.GetSheet(sheetName) != null)
                throw new InvalidOperationException($"A sheet with name '{sheetName}' already exists.");

            var newSheet = wb.CreateSheet(sheetName);
            if (position.HasValue && position.Value > 0)
            {
                int pos = position.Value - 1;
                if (pos < 0 || pos > wb.NumberOfSheets - 1)
                    pos = wb.NumberOfSheets - 1;

                wb.SetSheetOrder(sheetName, pos);
            }

            FlushWorkbook(wb);
        }

        public override void RenameSheet(string fromSheetName, string toSheetName)
        {
            using var wb = GetWorkbook();
            var sheet = wb.GetSheet(fromSheetName) ?? throw new InvalidOperationException($"No sheet with name '{fromSheetName}' was found.");
            if (sheet.SheetName == toSheetName)
                return;

            var anotherSheet = wb.GetSheet(toSheetName);
            if (anotherSheet is not null && wb.GetSheetIndex(sheet) != wb.GetSheetIndex(anotherSheet))
                throw new InvalidOperationException($"Another sheet with name '{toSheetName}' already exists in the workbook.");

            int sheetIndex = wb.GetSheetIndex(sheet);
            wb.SetSheetName(sheetIndex, toSheetName);
            FlushWorkbook(wb);
        }

        public override void FreezePanes(string sheetName, int colsToFreeze, int rowsToFreeze)
        {
            ValidateSheetName(sheetName);
            using var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet is null)
                return;

            sheet.CreateFreezePane(colsToFreeze, rowsToFreeze);

            FlushWorkbook(wb);
        }

        private void ToggleSheetState(string sheetName, SheetVisibility state)
        {
            ValidateSheetName(sheetName);
            using var wb = GetWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet is null)
                return;

            var index = wb.GetSheetIndex(sheet);
            wb.SetSheetVisibility(index, state);

            FlushWorkbook(wb);
        }

        public override void HideSheet(string sheetName) => 
            ToggleSheetState(sheetName, SheetVisibility.Hidden);

        public override void UnhideSheet(string sheetName) => 
            ToggleSheetState(sheetName, SheetVisibility.Visible);
    }
}