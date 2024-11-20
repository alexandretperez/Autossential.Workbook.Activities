using Autossential.Workbook.Core.Enums;
using System;
using System.Data;
using System.Drawing;

namespace Autossential.Workbook.Core.Processors
{
    public interface IWorkbookAdapter : IDisposable
    {
        void ActivateSheet(object sheetNameOrIndex);

        void AddHyperLink(string sheetName, string cell, string label, string link, string tooltip);

        void CreateNew();

        void Dispose(bool disposing);

        void FreezePanes(string sheetName, int cols, int rows);

        string[] GetHyperlinks(string sheetName, string range);

        string[] GetSheetNames();

        DataTable ReadRange(string sheetName, string range, bool addHeaders);

        int RemoveHyperLinks(string sheetName, string range);

        void Save();

        void WriteCell(string sheetName, string cell, object value);

        void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders);

        void AppendRange(string sheetName, DataTable dataTable);

        void FillColor(string sheetName, string range, Color[] colors, FillPattern pattern);

        void DrawBorder(string sheetName, string range, Border border, BorderStyle style, Color color);

        void RenameSheet(int sheetIndex, string newName);

        void DeleteSheet(string sheetName);

        void MergeRange(string sheetName, string range);

        void MoveSheet(string sheetName, int index, bool makeACopy = false, string copySheetName = null);
    }
}