using Microsoft.Win32.SafeHandles;
using System.IO;

namespace Autossential.Workbook.Core.Internals
{
    public sealed class WorkbookFileStream : FileStream
    {
        public WorkbookFileStream(SafeFileHandle handle, FileAccess access) : base(handle, access)
        {
        }

        public override void Close()
        {
            base.Close();
        }
    }
}
