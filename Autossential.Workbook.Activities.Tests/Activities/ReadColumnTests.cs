using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadColumnTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "A1", 0, 10)]
        [Arguments(".xls", "B1", 0, 0)]
        [Arguments(".xls", "F4", 0, 6)]
        [Arguments(".xls", "I3", 5, 5)]

        [Arguments(".xlsx", "A1", 0, 10)]
        [Arguments(".xlsx", "B1", 0, 0)]
        [Arguments(".xlsx", "F4", 0, 6)]
        [Arguments(".xlsx", "I3", 5, 5)]
        public async Task ReadColumn_ReturnsExpectedValue_BasedOnStaringCellAndLimit(string extension, string startingCell, int limit, int expectedCount)
        {
            object[] values = Run(extension, startingCell, limit);

            await Assert.That(values.Length).IsEqualTo(expectedCount);
        }

        [Test]

        [Arguments(".xls", null)]
        [Arguments(".xls", "")]

        [Arguments(".xlsx", null)]
        [Arguments(".xlsx", "")]
        public void ReadColumn_Fails_WhenMissingStartingCell(string extension, string? startingCell)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() => Run(extension, startingCell, 0));
        }

        [Test]
        public void ReadColumn_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new ReadColumn
                {
                    SheetName = "",
                    StartingCell = ""
                });
            });
        }

        private object[] Run(string extension, string? startingCell, int limit)
        {
            var data = TableUtils.Build(10, 10, (col, row) =>
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

            var values = InvokeWorkbookScopeWith(filePath, new ReadColumn
            {
                SheetName = "Sheet1",
                StartingCell = startingCell,
                Limit = limit
            });
            return values;
        }
    }
}