namespace Autossential.Workbook.Core.Internals
{
    internal class BIFF8CellReference : CellReference
    {
        private const uint MAX_ROWS = 65_536;
        private const uint MAX_COLS = 256; // IV (9 * 26 + 22)

        public BIFF8CellReference(uint row, uint col) : base(row, col) { }

        public BIFF8CellReference(string address) : base(address) { }

        protected override void ValidateConstraints()
        {
            if (Row > MAX_ROWS || Col > MAX_COLS)
            {
                Row = Col = 0;
            }
        }
        public static BIFF8CellReference Max() => new(MAX_ROWS, MAX_COLS);
    }
}
