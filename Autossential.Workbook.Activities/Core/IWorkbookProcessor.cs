using System.Data;

namespace Autossential.Workbook.Activities.Core
{
    public interface IWorkbookProcessor : IDisposable
    {
        void AppendRange(string sheetName, DataTable data);

        void DeleteSheet(string sheetName);

        (string, int, int) FindValue(string sheetName, string range, object value);

        void FreezePanes(string sheetName, int colsToFreeze, int rowsToFreeze);

        int GetColumnCount(string sheetName, string range);

        int GetRowCount(string sheetName, string range);

        string[] GetSheetNames();

        void HideSheet(string sheetName);

        void InsertSheet(string sheetName, int? position = null);

        object ReadCell(string sheetName, string address);

        object[] ReadColumn(string sheetName, string startingCell, int limit = 0);

        DataTable ReadRange(string sheetName, string range, bool hasHeaders, int headerRows = 1, int rowsPerRecord = 1);

        object[] ReadRow(string sheetName, string startingCell, int limit = 0);

        void RenameSheet(string fromSheetName, string toSheetName);

        void Save();

        void UnhideSheet(string sheetName);

        void WriteCell(string sheetName, string address, object value);

        void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders);

        void DeleteColumns(string sheetName, string references);
        void DeleteRows(string sheetName, string references);
    }
}