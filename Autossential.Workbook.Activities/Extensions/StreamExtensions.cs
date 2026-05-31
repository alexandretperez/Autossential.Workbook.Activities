namespace Autossential.Workbook.Activities.Extensions
{
    internal static class StreamExtensions
    {
        extension(Stream stream)
        {
            public Stream Reset()
            {
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }
        }
    }
}
