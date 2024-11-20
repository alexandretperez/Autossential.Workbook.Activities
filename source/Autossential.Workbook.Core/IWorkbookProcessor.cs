using System;
using System.Collections.Generic;
using System.Data;

namespace Autossential.Workbook.Core
{
    public interface IWorkbookProcessor : IDisposable
    {
        string[] GetSheetNames();
        int GetRowCount(string sheetName, string range);
        int GetColumnCount(string sheetName, string range);
        DataTable ReadRange(string sheetName, string range, bool hasHeaders);
    }
}
