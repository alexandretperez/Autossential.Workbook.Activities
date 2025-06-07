using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;

namespace Autossential.Workbook.Core.Extensions
{
    internal static class DataTableExtensions
    {
        public static void AddColumns(this DataTable dt, int startCol, int endCol, Func<int, int, string> columnNameHandler, ReadOnlyCollection<DbColumn> columnSchema)
        {
            for (int i = 0; i <= endCol - startCol; i++)
            {
                var colIndex = i + startCol - 1;
                if (colIndex < columnSchema.Count)
                {
                    dt.Columns.Add(columnNameHandler(colIndex, i), columnSchema[colIndex].DataType);
                    continue;
                }

                dt.Columns.Add(columnNameHandler(colIndex, i), typeof(object));
            }
        }
    }
}
