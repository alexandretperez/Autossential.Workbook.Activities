namespace Autossential.Workbook.Activities.Core
{
    internal record class BIFF8RangeReference : RangeReference
    {
        public BIFF8RangeReference(string range) : base(range, BIFF8CellReference.Max())
        {
        }

        public BIFF8RangeReference(CellReference start, CellReference end) : base(start, end)
        {
        }

        protected override CellReference NewCell(string address, int maxRows) => new BIFF8CellReference(address, maxRows);
    }
}
