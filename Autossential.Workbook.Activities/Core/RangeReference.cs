using System.Activities.Expressions;

namespace Autossential.Workbook.Activities.Core
{
    internal readonly struct RangeReference : IEquatable<RangeReference>
    {
        private RangeReference(CellReference start, CellReference end, RangeOrigin origin)
        {
            if (end.Col < start.Col || end.Row < start.Row)
                throw new ArgumentException("Range end must be greater than or equal to range start.");

            Start = start;
            End = end;
            Origin = origin;
        }

        public CellReference Start { get; }
        public CellReference End { get; }
        public RangeOrigin Origin { get; }
        public bool IsValid => Start.IsValid && End.IsValid;

        public static RangeReference OpenXml(int startCol, int startRow, int endCol, int endRow) =>
            new(CellReference.OpenXml(startCol, startRow), CellReference.OpenXml(endCol, endRow), RangeOrigin.Explicit);

        public static RangeReference Binary(int startCol, int startRow, int endCol, int endRow) =>
            new(CellReference.Binary(startCol, startRow), CellReference.Binary(endCol, endRow), RangeOrigin.Explicit);

        private static (string start, string end, RangeOrigin origin) Parse(string range, string maxReference)
        {
            if (string.IsNullOrEmpty(range))
                return (CellReference.MIN_REFERENCE, maxReference, RangeOrigin.Inferred);

            var sepIndex = range.IndexOf(':');
            if (sepIndex < 0)
                return (range, maxReference, RangeOrigin.InferredEnd);

            if (sepIndex == 0 || sepIndex == range.Length - 1 || range.IndexOf(':', sepIndex + 1) >= 0)
                throw new ArgumentException("The range address is invalid. The expected format is <start>:<end>, e.g.: A1:E9", nameof(range));

            return (range[..sepIndex], range[(sepIndex + 1)..], RangeOrigin.Explicit);
        }

        public static RangeReference OpenXml(string range)
        {
            var (start, end, origin) = Parse(range, CellReference.OPENXML_MAX_REFERENCE);
            return new RangeReference(
                CellReference.OpenXml(start),
                CellReference.OpenXml(end),
                origin
            );
        }

        public static RangeReference Binary(string range)
        {
            var (start, end, origin) = Parse(range, CellReference.BIFF8_MAX_REFERENCE);
            return new RangeReference(
                CellReference.Binary(start),
                CellReference.Binary(end),
                origin
            );
        }

        public readonly bool Equals(RangeReference other) =>
            Start == other.Start && End == other.End && Origin == other.Origin;

        public override readonly bool Equals(object obj) =>
            obj is RangeReference other && Equals(other);

        public static bool operator ==(RangeReference left, RangeReference right) =>
            left.Equals(right);

        public static bool operator !=(RangeReference left, RangeReference right) =>
            !left.Equals(right);

        public override readonly int GetHashCode() =>
            HashCode.Combine(Start, End, Origin);

        /// <summary>
        /// Determines if this range is equivalent to another range, ignoring the origin. This is used to determine if two ranges refer to the same cells, even if they were created differently (e.g. one with an explicit end and one with an inferred end).
        /// </summary>
        /// <param name="other">The other range to compare with.</param>
        /// <returns>True if the ranges are equivalent, false otherwise.</returns>
        public bool IsEquivalentTo(RangeReference other) => Start == other.Start && End == other.End;

        public bool IsRowInRange(int row) => row >= Start.Row && row <= End.Row;
        public bool IsColInRange(int col) => col >= Start.Col && col <= End.Col;

        public override readonly string ToString() => IsValid ? $"{Start}:{End}" : string.Empty;
    }
}
