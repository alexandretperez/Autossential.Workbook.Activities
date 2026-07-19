namespace Autossential.Workbook.Activities.Extensions
{
    internal static class EnumerableExtensions
    {
        extension(Enumerable)
        {
            public static IEnumerable<uint> Range(uint start, uint count)
            {
                for (uint i = start; i < start + count; i++)
                    yield return i;
            }
        }
    }
}
