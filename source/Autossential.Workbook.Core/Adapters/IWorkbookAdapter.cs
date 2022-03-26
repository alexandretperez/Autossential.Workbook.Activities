using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Autossential.Workbook.Core.Adapters
{
    public interface IWorkbookAdapter : IDisposable
    {
        Task<string[]> GetSheetNamesAsync();

        Task<bool> AddHyperLinkAsync(string sheetName, string cellAddress, string label, string address, string tooltip);

        Task<int> RemoveHyperlinksAsync(string sheetName, string range);

        Task<string[]> GetHyperlinksAsync(string sheetName, string range);

        Task<DataTable> ReadRangeAsync(string sheetName, string range, bool addHeaders, string password);

        Task SaveAsync();

        void CreateNew();

        void Dispose(bool disposing);
    }
}
