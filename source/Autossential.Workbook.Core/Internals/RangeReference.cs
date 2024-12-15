using NPOI.SS.Formula.Functions;

namespace Autossential.Workbook.Core.Internals
{
    internal abstract class RangeReference
    {
        public CellReference Start { get; protected set; }
        public CellReference End { get; protected set; }

        protected RangeReference(CellReference start, CellReference end)
        {
            Start = start;
            End = end;
            ValidateConstraints();
        }

        protected RangeReference(string range)
        {
            ParseRange(range);
            ValidateConstraints();
        }
        public bool IsValid => Start?.IsValid == true && End?.IsValid == true;
        protected abstract CellReference CreateCellReference(string address);
        protected abstract CellReference CreateMaxCellReference();
        protected void ParseRange(string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                Start = End = null;
                return;
            }

            string[] parts = range.Split(':');
            if (parts.Length > 2)
            {
                Start = End = null;
                return;
            }

            Start = CreateCellReference(parts[0]);
            End = parts.Length > 1
                ? CreateCellReference(parts[1])
                : CreateMaxCellReference();
        }
        protected void ValidateConstraints()
        {
            if (Start.Row > End.Row || Start.Col > End.Col)
                Start = End = null;
        }

        public string GetRange()
        {
            if (!IsValid) return string.Empty;
            return $"{Start.ToString()}:{End.ToString()}";
        }

        public bool IsInRange(uint row, uint col) =>
            IsRowInRange(row) && IsColInRange(col);

        public bool IsRowInRange(uint row) =>
            row >= Start.Row &&
            row <= End.Row;

        public bool IsColInRange(uint col) =>
            col >= Start.Col &&
            col <= End.Col;

        public override string ToString() => $"{GetRange()} (Start [row: {Start.Row}, col: {Start.Col}], End [row: {End.Row}, col: {End.Col}])";
    }

}
