using Autossential.Workbook.Activities.Core;
using System.Data;

namespace Autossential.Workbook.Activities.Extensions
{
    internal static class DataTableExtensions
    {
        extension(DataTable table)
        {
            public DataTable TrimOrAppend(RangeReference range, string colNamePrefix, int colNameIndex, bool hasHeaders, int headerRows, int rowsPerRecord)
            {
                static bool HasValue(DataRow row, int columnIndex)
                {
                    if (row.IsNull(columnIndex))
                        return false;

                    return row[columnIndex]?.ToString().Length > 0;
                }

                int lastNonEmptyRowIndex = -1;
                int lastNonEmptyColIndex = -1;

                // last non-empty row
                for (int ri = table.Rows.Count - 1; ri >= 0; ri--)
                {
                    var row = table.Rows[ri];

                    for (int ci = 0; ci < table.Columns.Count; ci++)
                    {
                        if (HasValue(row, ci))
                        {
                            lastNonEmptyRowIndex = ri;
                            break;
                        }
                    }

                    if (lastNonEmptyRowIndex >= 0)
                        break;
                }

                // last non-empty column
                for (int ci = table.Columns.Count - 1; ci >= 0; ci--)
                {
                    for (int ri = 0; ri < table.Rows.Count; ri++)
                    {
                        var row = table.Rows[ri];

                        if (HasValue(row, ci))
                        {
                            lastNonEmptyColIndex = ci;
                            break;
                        }
                    }

                    if (lastNonEmptyColIndex >= 0)
                        break;
                }

                if (range.InputType == RangeInputType.A1B1 || range.InputType == RangeInputType.AB1)
                {
                    int expectedRowCount = range.End.Row;
                    expectedRowCount -= (range.Start.Row - 1);

                    if (hasHeaders)
                        expectedRowCount -= headerRows;

                    if (rowsPerRecord > 1)
                        expectedRowCount = (int)(Math.Ceiling(expectedRowCount / (double)rowsPerRecord));

                    while (table.Rows.Count < expectedRowCount)
                        table.Rows.Add(table.NewRow());

                    int expectedColumnCount = range.End.Col;
                    expectedColumnCount -= (range.Start.Col - 1);

                    while (table.Columns.Count < expectedColumnCount)
                    {
                        string name;
                        do
                        {
                            name = $"{colNamePrefix}{colNameIndex++}";
                        } while (table.Columns.Contains(name));

                        table.Columns.Add(name, typeof(object));
                    }
                }
                else
                {
                    for (int i = table.Columns.Count - 1; i > lastNonEmptyColIndex; i--)
                        table.Columns.RemoveAt(i);

                    for (int i = table.Rows.Count - 1; i > lastNonEmptyRowIndex; i--)
                        table.Rows.RemoveAt(i);
                }

                return table;
            }
        }
    }
}
