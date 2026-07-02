namespace Autossential.Workbook.Activities.Core
{
    internal abstract record class RangeReference
    {
        public CellReference Start { get; private set; }
        public CellReference End { get; private set; }
        protected RangeReference(string range, CellReference max)
        {
            if (string.IsNullOrWhiteSpace(range))
                throw new ArgumentNullException(nameof(range), "The 'range' argument cannot be null or empty");

            var sepIndex = range.IndexOf(':');
            if (sepIndex < 0)
            {
                Start = NewCell(range, 1);
                End = max;
                InputType = Start.IsRowInferred() ? RangeInputType.A : RangeInputType.A1;
                return;
            }

            if (sepIndex == 0 || sepIndex == range.Length - 1 || range.IndexOf(':', sepIndex + 1) >= 0)
                throw new ArgumentException("The range address is invalid. The expected format is <start> or <start>:<end>, e.g.: A1 or A1:E9", nameof(range));


            Start = NewCell(range[..sepIndex], 1);
            End = NewCell(range[(sepIndex + 1)..], max.Row);
            Validate(Start, End);

            if (Start.IsRowInferred() && End.IsRowInferred())
            {
                InputType = RangeInputType.AB;
            }
            else if (Start.IsRowInferred())
            {
                InputType = RangeInputType.AB1;
            }
            else if (End.IsRowInferred())
            {
                InputType = RangeInputType.A1B;
            }
            else
            {
                InputType = RangeInputType.A1B1;
            }
        }
        protected RangeReference(CellReference start, CellReference end)
        {
            Validate(start, end);
            Start = start;
            End = end;
        }

        private static void Validate(CellReference start, CellReference end)
        {
            if (start > end)
                throw new ArgumentException($"Invalid range: the start cell '{start}' cannot be greater than the end cell '{end}'.");
        }
        public RangeInputType InputType { get; private set; }
        protected abstract CellReference NewCell(string address, int maxRows);
        public override int GetHashCode() => HashCode.Combine(Start, End);
    }
}
