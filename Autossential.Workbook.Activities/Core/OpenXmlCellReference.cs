namespace Autossential.Workbook.Activities.Core
{
    internal record class OpenXmlCellReference : CellReference
    {
        protected const int MaxRows = 1_048_576;
        protected const int MaxColumns = 16_384;

        public OpenXmlCellReference(string address) : base(address, MaxRows)
        {
        }

        public OpenXmlCellReference(int col, int row) : base(col, row)
        {
        }

        public OpenXmlCellReference(string address, int maxRows) : base(address, maxRows)
        {
        }

        public static OpenXmlCellReference Max() => new(MaxColumns, MaxRows);

        protected override (int col, int row) MaxReference() => (MaxColumns, MaxRows);
    }
}