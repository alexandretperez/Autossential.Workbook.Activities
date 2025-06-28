using System;
using System.IO;

namespace Autossential.Workbook.Core.Extensions
{
    internal static class StreamExtensions
    {
        public static int CalculateBufferSize(this Stream stream)
        {
            if (stream == null || !stream.CanRead)
                throw new ArgumentNullException(nameof(stream), "Stream cannot be null and must be readable.");

            long len = stream.Length;
            const int defaultBufferSize = 81920;

            if (len < defaultBufferSize)
                return (int)len;

            if (len < 10 * 1024 * 1024) // < 10 MB
                return defaultBufferSize;

            if (len < 100 * 1024 * 1024) // < 100 MB
                return 256 * 1024;

            return 1024 * 1024; // 1 MB for larger files
        }
    }
}
