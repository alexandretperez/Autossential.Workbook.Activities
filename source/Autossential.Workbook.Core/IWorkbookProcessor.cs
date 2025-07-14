using System;
using System.Data;

namespace Autossential.Workbook.Core
{
    public interface IWorkbookProcessor : IDisposable
    {
        string[] GetSheetNames();
        int GetRowCount(string sheetName, string range);
        int GetColumnCount(string sheetName, string range);
        DataTable ReadRange(string sheetName, string range, bool hasHeaders, bool useColumnDataType);
        void RenameSheet(int sheetIndex, string newSheetName);
        void RenameSheet(string fromSheetName, string toSheetName);
        void DeleteSheet(string sheetName);
        void ActivateSheet(string sheetName);
        void ActivateSheet(int sheetIndex);
        (int index, string name) GetActiveSheet();
        void Save();
        void WriteRange(DataTable dt, string sheetName, string startCell, bool addHeaders);
    }
}
