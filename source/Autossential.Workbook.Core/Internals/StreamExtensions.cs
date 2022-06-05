using System.IO;

namespace Autossential.Workbook.Core.Internals
{
    public static class StreamExtensions
    {
        public static Stream Reset(this Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }
    }
}
