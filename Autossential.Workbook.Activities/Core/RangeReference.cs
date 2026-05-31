namespace Autossential.Workbook.Activities.Core
{
    internal class RangeReference : IEquatable<RangeReference>
    {
        public RangeReference(CellReference start, CellReference end)
        {
            Start = start;
            End = end;
            Origin = RangeOrigin.Explicit;
        }

        public RangeReference(ExcelFileType type, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                Start = new CellReference(type, "A1");
                End = CellReference.Max(type);
                Origin = RangeOrigin.Default;
                return;
            }

            var parts = range.Split(':');
            if (parts.Length == 1)
            {
                Start = new CellReference(type, parts[0]);
                End = CellReference.Max(type);
                Origin = RangeOrigin.InferredEnd;
            }
            else if (parts.Length == 2)
            {
                Start = new CellReference(type, parts[0]);
                End = new CellReference(type, parts[1]);
                Origin = RangeOrigin.Explicit;
            }
            else
            {
                throw new ArgumentException("Invalid range address format.", nameof(range));
            }
        }

        public CellReference Start { get; }
        public CellReference End { get; }
        public RangeOrigin Origin { get; }

        public override string ToString()
        {
            return $"{Start}:{End}";
        }

        public bool IsRowInRange(int row) => row >= Start.Row && row <= End.Row;
        public bool IsColInRange(int col) => col >= Start.Col && col <= End.Col;

        public bool Equals(RangeReference other)
        {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            return Start.Equals(other.Start) && End.Equals(other.End);
        }

        public override bool Equals(object obj) => Equals(obj as RangeReference);

        public override int GetHashCode() => HashCode.Combine(Start, End);

        public static bool operator ==(RangeReference left, RangeReference right)
            => EqualityComparer<RangeReference>.Default.Equals(left, right);

        public static bool operator !=(RangeReference left, RangeReference right)
            => !(left == right);
    }
}
