using System.Data;

namespace Autossential.Workbook.Activities.Core
{
    public interface IWorkbookProcessor : IDisposable
    {
        string[] GetSheetNames();
        DataTable ReadRange(string sheetName, string range, bool hasHeaders, int headerRows = 1, int rowsPerRecord = 1);
        int GetRowCount(string sheetName, string range);
        int GetColumnCount(string sheetName, string range);
        object ReadCell(string sheetName, string address);
        object[] ReadRow(string sheetName, string startingCell, int limit = 0);
        object[] ReadColumn(string sheetName, string startingCell, int limit = 0);
        void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders);
        void WriteCell(string sheetName, string address, object value);
        (string, int, int) FindValue(string sheetName, string range, object value);
        void DeleteSheet(string sheetName);
        void InsertSheet(string sheetName, int? position = null);
        void RenameSheet(string fromSheetName, string toSheetName);
        void FreezePanes(string sheetName, int colsToFreeze, int rowsToFreeze);
        void HideSheet(string sheetName);
        void UnhideSheet(string sheetName);
        void Save();
    }
}
