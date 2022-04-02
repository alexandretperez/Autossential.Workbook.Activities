using System;
using System.Data;
using System.Threading.Tasks;

namespace Autossential.Workbook.Core.Adapters
{
    public interface IWorkbookAdapter : IDisposable
    {
        Task<string[]> GetSheetNamesAsync();

        Task<bool> AddHyperLinkAsync(string sheetName, string cellAddress, string label, string address, string tooltip);

        Task<int> RemoveHyperlinksAsync(string sheetName, string range);

        Task<string[]> GetHyperlinksAsync(string sheetName, string range);

        Task<DataTable> ReadRangeAsync(string sheetName, string range, bool addHeaders);

        Task WriteCellAsync(string sheetName, string cellAddress, object value);

        Task WriteRangeAsync(string sheetName, string cellAddress, DataTable value, bool addHeaders);

        void Save();

        void CreateNew();

        void Dispose(bool disposing);
    }
}
