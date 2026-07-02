using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class GetRangeSizeTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "", 10, 10)]
        [Arguments(".xls", null, 10, 10)]

        [Arguments(".xlsx", "", 10, 10)]
        [Arguments(".xlsx", null, 10, 10)]
        public async Task GetRangeSize_ReturnsBasedOnWholeSheet_WhenMissingRange(string extension, string? range, int expectedCols, int expectedRows)
        {
            Tuple<int, int> result = Run(extension, range);

            await Assert.That(result.Item1).IsEqualTo(expectedCols);
            await Assert.That(result.Item2).IsEqualTo(expectedRows);
        }

        [Test]

        [Arguments(".xls", "A1", 10, 10)]
        [Arguments(".xls", "B1:G10", 5, 10)]
        [Arguments(".xls", "G6:H10", 0, 0)]
        [Arguments(".xls", "F7:G10", 1, 3)]

        [Arguments(".xlsx", "A1", 10, 10)]
        [Arguments(".xlsx", "B1:G10", 5, 10)]
        [Arguments(".xlsx", "G6:H10", 0, 0)]
        [Arguments(".xlsx", "F7:G10", 1, 3)]
        public async Task GetRangeSize_ReturnsExpectedSize_BasedOnRange(string extension, string range, int expectedCols, int expectedRows)
        {
            Tuple<int, int> result = Run(extension, range);

            await Assert.That(result.Item1).IsEqualTo(expectedCols);
            await Assert.That(result.Item2).IsEqualTo(expectedRows);
        }

        [Test]
        [Arguments(".xlsx", "A1:C", 3, 3)]
        [Arguments(".xlsx", "B:D5", 3, 4)]
        [Arguments(".xlsx", "E3:J", 6, 8)]
        [Arguments(".xlsx", "A:H", 8, 8)]

        [Arguments(".xls", "A1:C", 3, 3)]
        [Arguments(".xls", "B:D5", 3, 4)]
        [Arguments(".xls", "E3:J", 6, 8)]
        [Arguments(".xls", "A:H", 8, 8)]
        public async Task GetRangeSize_ReturnsExpectedSize_BasedOnInferredRange(string extension, string range, int expectedCols, int expectedRows)
        {
            var data = TableUtils.Build(10, 10, (col, row) =>
            {
                if (col == row) return $"C{col}R{row}";
                return "";
            });

            Tuple<int, int> result = Run(extension, range, data);

            await Assert.That(result.Item1).IsEqualTo(expectedCols);
            await Assert.That(result.Item2).IsEqualTo(expectedRows);
        }

        [Test]
        public void GetRangeSize_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new GetRangeSize
                {
                    SheetName = ""
                });
            });
        }

        private Tuple<int, int> Run(string extension, string? range, DataTable? customTable = null)
        {
            var data = customTable ?? TableUtils.Build(10, 10, (col, row) =>
            {
                var value = $"C{col}R{row}";
                return col switch
                {
                    2 or 7 => null,
                    3 => row % 2 == 0 || row % 7 == 0 ? value : "",
                    8 => row < 6 ? value : null,
                    4 or 5 or 6 => (row + col) % 3 == 0 ? value : DBNull.Value,
                    10 => row > 4 && row < 10 ? value : "",
                    _ => value
                };
            });

            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            processor.Save();

            var rowCount = new Variable<int>();
            var colCount = new Variable<int>();

            var result = InvokeWorkbookScopeWith(filePath, [rowCount, colCount], [new GetRangeSize
                    {
                        SheetName = "Sheet1",
                        Range = range,
                        RowCount = new OutArgument<int>(rowCount),
                        ColumnCount = new OutArgument<int>(colCount)
                    }], env => Tuple.Create(colCount.Get(env), rowCount.Get(env)));
            return result;
        }
    }
}