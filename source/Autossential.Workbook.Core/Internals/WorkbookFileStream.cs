using System.IO;

namespace Autossential.Workbook.Core.Internals
{
    public sealed class WorkbookFileStream : FileStream
    {
        public WorkbookFileStream(string path, FileMode mode) : base(path, mode)
        {
        }

        public override void Close()
        {
            base.Close();
        }

        public void CloseWorkbook()
        {
            base.Close();
        }
    }

}
