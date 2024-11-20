namespace Autossential.Workbook.Core.Processors
{
    public class OpenXMLWorkbookProcessor : WorkbookProcessorBase
    {
        public OpenXMLWorkbookProcessor(string filePath, bool createIfNotExist = true) : base(filePath, createIfNotExist)
        {
        }

        public override void CreateNew()
        {
            // TODO
        }
    }
}