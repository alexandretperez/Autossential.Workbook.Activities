using System.Data;

namespace Autossential.Workbook.Activities.Extensions
{
    internal static class DataTableExtensions
    {
        extension(DataTable table)
        {
            public void RemoveTrailingRows(int fromRowIndex = 0)
            {
                for (int i = table.Rows.Count - 1; i >= fromRowIndex; i--)
                {
                    var row = table.Rows[i];
                    if (row.ItemArray.All(value => value == null || string.IsNullOrEmpty(value.ToString())))
                    {
                        table.Rows.RemoveAt(i);
                        continue;
                    }

                    break;
                }
            }

            public void RemoveTrailingColumns(int fromColumnIndex = 0)
            {
                int index = table.Columns.Count;
                while (--index > fromColumnIndex)
                {
                    var isEmpty = true;
                    foreach (DataRow row in table.Rows)
                    {
                        var value = row[index];
                        if (!row.IsNull(index) && value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        table.Columns.RemoveAt(index);
                        continue;
                    }

                    break;
                }
            }

            public void AddTrailingRows(int rows)
            {
                for (int i = table.Rows.Count; i < rows; i++)
                    table.Rows.Add(table.NewRow());
            }

            public void AddTrailingColumns(int cols, int columnIndex, string columnPrefix)
            {
                for (int i = table.Columns.Count; i < cols; i++)
                    table.Columns.Add($"{columnPrefix}{columnIndex++}", typeof(object));
            }
        }
    }
}
