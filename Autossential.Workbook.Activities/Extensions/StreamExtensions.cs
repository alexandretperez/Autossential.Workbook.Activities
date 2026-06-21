using System.Security.Cryptography;

namespace Autossential.Workbook.Activities.Extensions
{
    internal static class StreamExtensions
    {
        extension(Stream stream)
        {
            /// <summary>
            /// Calculates an optimal buffer size for reading the stream based on its length. For small streams, it returns the stream's length as the buffer size. For larger streams, it returns a default buffer size of 81920 bytes (80 KB) or larger depending on the stream's length, up to a maximum of 1 MB for very large streams.
            /// </summary>
            /// <returns>The optimal buffer size.</returns>
            /// <exception cref="ArgumentNullException"></exception>
            public int CalculateBufferSize()
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

            /// <summary>
            /// Computes a SHA256 hash of the stream's contents. The stream's position is reset to the beginning before hashing.
            /// </summary>
            /// <returns>A string representation of the SHA256 hash.</returns>
            public string ComputeHash()
            {
                stream.Position = 0;
                using var sha256 = SHA256.Create();
                var bytes = sha256.ComputeHash(stream);
                return BitConverter.ToString(bytes);
            }
        }
    }
}
