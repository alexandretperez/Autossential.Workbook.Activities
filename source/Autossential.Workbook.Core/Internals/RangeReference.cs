using System;

namespace Autossential.Workbook.Core.Internals
{
    internal struct RangeReference
    {
        public CellReference Start { get; private set; }
        public CellReference End { get; private set; }

        public RangeReference()
        {
            Start = new CellReference();
            End = new CellReference();
        }

        private RangeReference(CellReference start, CellReference end)
        {
            Start = start;
            End = end;
        }

        private static RangeReference Create(string range, Func<string, CellReference> cellRefFn, Func<CellReference> maxCellRefFn)
        {
            if (string.IsNullOrEmpty(range))
                return new RangeReference();

            var parts = range.Split(':');
            if (parts.Length > 2)
                return new RangeReference();

            var start = cellRefFn(parts[0]);
            var end = parts.Length == 2 ? cellRefFn(parts[1]) : maxCellRefFn();
            return new RangeReference(start, end);
        }

        public static RangeReference CreateForOpenXml(string range)
            => Create(range, CellReference.CreateForOpenXml, CellReference.GetMaxForOpenXml);

        public static RangeReference CreateForBIFF8(string range)
            => Create(range, CellReference.CreateForBIFF8, CellReference.GetMaxForBIFF8);

        public readonly bool IsValid =>
            Start.IsValid
            && End.IsValid
            && Start.Row < End.Row
            && Start.Col < End.Col;

        public override string ToString()
        {
            return $"{Start}:{End}";
        }

        public readonly bool IsRowInRange(long row) => row >= Start.Row && row <= End.Row;
        public readonly bool IsColInRange(long col) => col >= Start.Col && col <= End.Col;
    }
}
