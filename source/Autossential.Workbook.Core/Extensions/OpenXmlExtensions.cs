using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Linq;

namespace Autossential.Workbook.Core.Extensions
{
    public static class OpenXmlExtensions
    {
        public static string GetCellValue(this Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.CellValue == null)
                return null;

            string value = cell.CellValue.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                return sharedStringTable.ElementAt(int.Parse(value)).InnerText;

            if (cell.StyleIndex != null)
            {
                // TODO: Add specific formatting logic here if needed
                // Handle formatted numbers, dates, etc.
            }

            return value;
        }
        public static WorksheetPart GetWorksheetPartByName(this WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
                return null;

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }


        public static CellValues GetEquivalentCellDataType(this object value)
        {
            if (value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal)
                return CellValues.Number;

            if (value is bool)
                return CellValues.Boolean;

            if (value is DateTime)
                return CellValues.Date;

            return CellValues.String;
        }
   }
}
