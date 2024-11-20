namespace Autossential.Workbook.Core.Internals
{
    internal class BIFF8RangeReference : RangeReference
    {
        public BIFF8RangeReference(BIFF8CellReference start, BIFF8CellReference end)
            : base(start, end) { }

        public BIFF8RangeReference(string range) : base(range) { }

        protected override CellReference CreateCellReference(string address) => new BIFF8CellReference(address);
        protected override CellReference CreateMaxCellReference() => BIFF8CellReference.Max();
    }

}
