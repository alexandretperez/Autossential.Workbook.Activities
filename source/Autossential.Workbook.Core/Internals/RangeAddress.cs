namespace Autossential.Workbook.Core.Internals
{
    internal class RangeAddress
    {
        public RangeAddress(string range, bool zeroBasedIndex = false)
        {
            if (string.IsNullOrEmpty(range))
            {
                First = CellAddress.UseDefault("A1", zeroBasedIndex);
                Last = CellAddress.UseDefault("B2", zeroBasedIndex);
                return;
            }

            var addresses = range.Split(':');

            First = new CellAddress(addresses[0], zeroBasedIndex);
            Last = new CellAddress(addresses.Length == 2 ? addresses[1] : null, zeroBasedIndex);
        }

        public CellAddress First { get; set; }
        public CellAddress Last { get; set; }
        public int RowsUsed() => Last.Row + 1 - First.Row;
        public int ColsUsed() => Last.Col + 1 - First.Col;
    }
}
