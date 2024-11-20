namespace Autossential.Workbook.Core.Internals
{
    internal class OpenXMLCellReference : CellReference
    {
        private const uint MAX_ROWS = 1_048_576;
        private const uint MAX_COLS = 16_384; // XFD (24 * 26 * 26 + 6 * 26 + 4)

        public OpenXMLCellReference(uint row, uint col) : base(row, col) { }

        public OpenXMLCellReference(string address) : base(address) { }

        protected override void ValidateConstraints()
        {
            if (Row > MAX_ROWS || Col > MAX_COLS)
            {
                Row = Col = 0;
            }
        }

        public static OpenXMLCellReference Max() => new(MAX_ROWS, MAX_COLS);
    }
}
