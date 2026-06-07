using System.Data;

namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class BinaryWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        public override void WriteRange(string sheetName, DataTable data, string startingCell, bool addHeaders)
        {
            throw new NotImplementedException();
        }

        protected override CellReference ResolveCell(string address) => CellReference.Binary(address);

        protected override RangeReference ResolveRange(string range) => RangeReference.Binary(range);
    }
}
