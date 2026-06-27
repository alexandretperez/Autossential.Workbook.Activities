namespace Autossential.Workbook.Activities.Core
{
    internal readonly struct RangeReference : IEquatable<RangeReference>
    {
        public RangeReference(string range, string maxReference)
        {
            if (string.IsNullOrWhiteSpace(range))
                throw new ArgumentNullException(nameof(range), "The 'range' argument cannot be null or empty");

            var sepIndex = range.IndexOf(':');
            if (sepIndex < 0)
            {
                Start = new CellReference(range);
                End = new CellReference(maxReference);
                InputType = Start.HasInferredRow() ? RangeInputType.A : RangeInputType.A1;
                return;
            }

            if (sepIndex == 0 || sepIndex == range.Length - 1 || range.IndexOf(':', sepIndex + 1) >= 0)
                throw new ArgumentException("The range address is invalid. The expected format is <start> or <start>:<end>, e.g.: A1 or A1:E9", nameof(range));

            Start = new CellReference(range[..sepIndex]);
            End = new CellReference(range[(sepIndex + 1)..]);

            if (Start > End)
                throw new ArgumentException($"Invalid range: the start cell '{Start}' cannot be greater than the end cell '{End}'.");

            if (Start.HasInferredRow() && End.HasInferredRow())
            {
                InputType = RangeInputType.AB;
            }
            else if (Start.HasInferredRow())
            {
                InputType = RangeInputType.AB1;
            }
            else if (End.HasInferredRow())
            {
                InputType = RangeInputType.A1B;
            }
            else
            {
                InputType = RangeInputType.A1B1;
            }
        }

        public CellReference Start { get; }
        public CellReference End { get; }
        public RangeInputType InputType { get; }
        public bool IsValidForOpenXml() => Start.IsValidForOpenXml() && End.IsValidForOpenXml();
        public bool IsValidForBIFF8() => Start.IsValidForBIFF8() && End.IsValidForBIFF8();

        public readonly bool Equals(RangeReference other) =>
            Start == other.Start && End == other.End && InputType == other.InputType;

        public override readonly bool Equals(object obj) =>
            obj is RangeReference other && Equals(other);

        public static bool operator ==(RangeReference left, RangeReference right) =>
            left.Equals(right);

        public static bool operator !=(RangeReference left, RangeReference right) =>
            !left.Equals(right);

        public override readonly int GetHashCode() =>
            HashCode.Combine(Start, End, InputType);

        public bool IsEquivalentTo(RangeReference other) => Start == other.Start && End == other.End;

        public override readonly string ToString() => $"{Start}:{End}";
    }
}
