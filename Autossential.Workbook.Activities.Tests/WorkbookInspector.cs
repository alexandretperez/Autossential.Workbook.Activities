using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Windows.Controls;

namespace Autossential.Workbook.Activities.Tests
{
    public static class WorkbookInspector
    {
        public record struct Info
        {
            public bool IsVisible { get; set; }
            public bool IsFrozen { get; set; }
            public int ColsFrozen { get; set; }
            public int RowsFrozen { get; set; }
        }

        public static Info Inspect(string workbookFilePath, string sheetName)
        {
            if (string.Equals(".xlsx", Path.GetExtension(workbookFilePath), StringComparison.OrdinalIgnoreCase))
            {
                using var doc = SpreadsheetDocument.Open(workbookFilePath, false);
                return Inspect(doc, sheetName);
            }
            else
            {
                using var fs = File.OpenRead(workbookFilePath);
                return Inspect(new HSSFWorkbook(fs), sheetName);
            }
        }

        public static Info Inspect(SpreadsheetDocument doc, string sheetName)
        {
            var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var visible = sheet.State is null || sheet.State.Value == SheetStateValues.Visible;
            var wsPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var pane = wsPart.Worksheet
                                   .GetFirstChild<SheetViews>()?
                                   .GetFirstChild<SheetView>()?
                                   .GetFirstChild<Pane>();

            if (pane is null || pane.State?.Value != PaneStateValues.Frozen)
            {
                return new Info
                {
                    IsVisible = visible,
                    IsFrozen = false,
                    ColsFrozen = 0,
                    RowsFrozen = 0
                };
            }

            return new Info
            {
                IsVisible = visible,
                IsFrozen = true,
                ColsFrozen = (int)(pane.HorizontalSplit?.Value ?? 0),
                RowsFrozen = (int)(pane.VerticalSplit?.Value ?? 0)
            };
        }

        public static Info Inspect(IWorkbook wb, string sheetName)
        {
            var sheetIndex = wb.GetSheetIndex(sheetName);
            if (sheetIndex < 0)
                throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

            var visible = !(wb.IsSheetHidden(sheetIndex) || wb.IsSheetVeryHidden(sheetIndex));
            var sheet = wb.GetSheetAt(sheetIndex);
            var pane = sheet.PaneInformation;

            if (pane is null || !pane.IsFreezePane())
            {
                return new Info
                {
                    IsVisible = visible,
                    IsFrozen = false,
                    ColsFrozen = 0,
                    RowsFrozen = 0
                };
            }

            return new Info
            {
                IsVisible = visible,
                IsFrozen = true,
                ColsFrozen = pane.VerticalSplitPosition,
                RowsFrozen = pane.HorizontalSplitPosition
            };
        }
    }
}