namespace Autossential.Workbook.Core.Internals
{
    internal class OpenXMLRangeReference : RangeReference
    {
        public OpenXMLRangeReference(OpenXMLCellReference start, OpenXMLCellReference end)
            : base(start, end) { }

        public OpenXMLRangeReference(string range) : base(range) { }

        protected override CellReference CreateCellReference(string address) => new OpenXMLCellReference(address);
        protected override CellReference CreateMaxCellReference() => OpenXMLCellReference.Max();
    }

}
