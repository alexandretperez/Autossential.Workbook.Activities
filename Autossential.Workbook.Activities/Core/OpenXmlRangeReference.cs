namespace Autossential.Workbook.Activities.Core
{
    internal record class OpenXmlRangeReference : RangeReference
    {
        public OpenXmlRangeReference(string range) : base(range, OpenXmlCellReference.Max())
        {
        }

        public OpenXmlRangeReference(CellReference start, CellReference end) : base(start, end)
        {
        }

        protected override CellReference NewCell(string address, int maxRows) => new OpenXmlCellReference(address, maxRows);
    }
}
