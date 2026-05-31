namespace Autossential.Workbook.Activities.Core.Processors
{
    internal class BinaryWorkbookProcessor(string filePath, string password) : WorkbookProcessorBase(filePath, password)
    {
        private const ExcelFileType FileType = ExcelFileType.Binary;

        protected override CellReference ResolveCell(string address)
        {
            return new CellReference(FileType, address);
        }

        protected override RangeReference ResolveRange(string range)
        {
            return new RangeReference(FileType, range);
        }
    }
}
