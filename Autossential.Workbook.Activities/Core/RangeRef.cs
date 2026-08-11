namespace Autossential.Workbook.Activities.Core
{
    internal readonly record struct RangeRef(CellRef Start, CellRef End)
    {
        public static RangeRef Parse(ReadOnlySpan<char> span)
        {
            int sep = span.IndexOf(':');
            if (sep == -1)
                return new RangeRef(CellRef.Parse(span), new CellRef());

            var start = span[..sep];
            var end = span[(sep + 1)..];
            return new RangeRef(CellRef.Parse(start), CellRef.Parse(end));
        }
        public RangeInputType GetInputType()
        {
            var sc = Start.Col > 0;
            var sr = Start.Row > 0;
            var ec = End.Col > 0;
            var er = End.Row > 0;

            if (sc && sr && ec && er) return RangeInputType.A1B1;
            if (sc && sr && ec) return RangeInputType.A1B;
            if (sc && ec && er) return RangeInputType.AB1;
            if (sc && sr) return RangeInputType.A1;
            if (sc && ec) return RangeInputType.AB;
            if (sc) return RangeInputType.A;
            return RangeInputType.None;
        }

        public RangeRef Normalize(CellRef maxRef)
        {
            return GetInputType() switch
            {
                RangeInputType.A1B => new RangeRef(Start, new CellRef(End.Col, maxRef.Row)),
                RangeInputType.AB1 => new RangeRef(new CellRef(Start.Col, 1), End),
                RangeInputType.A1 => new RangeRef(Start, maxRef),
                RangeInputType.AB => new RangeRef(new CellRef(Start.Col, 1), new CellRef(End.Col, maxRef.Row)),
                RangeInputType.A => new RangeRef(new CellRef(Start.Col, 1), maxRef),              
                _ => this
            };
        }

        public static RangeRef MaxOpenXML() => new(new CellRef(1, 1), CellRef.MaxOpenXML());
        public static RangeRef MaxBIFF8() => new(new CellRef(1, 1), CellRef.MaxBIFF8());
    }
}