using NPOI.SS.UserModel;

namespace Autossential.Workbook.Core.Internals
{
    public static class SheetExtensions
    {
        public static IRow Row(this ISheet sheet, int row)
        {
            return sheet.GetRow(row) ?? sheet.CreateRow(row);
        }

        public static ICell Cell(this IRow row, int cell)
        {
            return row.GetCell(cell) ?? row.CreateCell(cell);
        }
    }
}
