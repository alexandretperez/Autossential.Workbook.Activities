namespace Autossential.Workbook.Activities.Core
{
    internal record class BIFF8CellReference : CellReference
    {
        protected const int MaxRows = 65_536;
        protected const int MaxColumns = 256;

        public BIFF8CellReference(string address) : base(address, MaxRows)
        {
        }

        public BIFF8CellReference(int col, int row) : base(col, row)
        {
        }

        public BIFF8CellReference(string address, int maxRows) : base(address, maxRows)
        {
        }

        public static BIFF8CellReference Max() => new(MaxColumns, MaxRows);

        protected override (int col, int row) MaxReference() => (MaxColumns, MaxRows);
    }
}