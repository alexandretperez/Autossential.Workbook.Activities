using System.Data;

namespace Autossential.Workbook.Activities.Core
{
    public interface IWorkbookProcessor : IDisposable
    {
        string[] GetSheetNames();
        DataTable ReadRange(string sheetName, string range, bool hasHeaders, int headerRows, int rowsPerRecord);
        int GetRowCount(string sheetName, string range);
        int GetColumnCount(string sheetName, string range);
        object ReadCell(string sheetName, string address);
        object[] ReadRow(string sheetName, string startingCell, int limit);
        object[] ReadColumn(string sheetName, string startingCell, int limit);
        void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders);
        void Save();
    }
}
