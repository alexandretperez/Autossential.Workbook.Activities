namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadRowTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "")]
        [Arguments(".xls", null)]

        [Arguments(".xlsx", "")]
        [Arguments(".xlsx", null)]
        public async Task ReadRow_Fails_WhenMissingStartingCell(string extension, string? startingCell)
        {
            await Assert.ThrowsAsync<InvalidOperationException>(() => Task.FromResult(Run(extension, startingCell, 0)));
        }

        [Test]

        [Arguments(".xls", "A1", 0, 10)]
        [Arguments(".xls", "A2", 0, 0)]
        [Arguments(".xls", "D6", 0, 6)]
        [Arguments(".xls", "C9", 5, 5)]

        [Arguments(".xlsx", "A1", 0, 10)]
        [Arguments(".xlsx", "A2", 0, 0)]
        [Arguments(".xlsx", "D6", 0, 6)]
        [Arguments(".xlsx", "C9", 5, 5)]
        public async Task ReadRow_ReturnsExpectedValue_BasedOnStaringCellAndLimit(string extension, string startingCell, int limit, int expectedCount)
        {
            object[] values = Run(extension, startingCell, limit);

            await Assert.That(values.Length).IsEqualTo(expectedCount);
        }

        private object[] Run(string extension, string? startingCell, int limit)
        {
            var data = TableUtils.Build(10, 10, (col, row) =>
            {
                var value = $"C{col}R{row}";
                return row switch
                {
                    2 or 7 => null,
                    3 => col % 2 == 0 || col % 7 == 0 ? value : "",
                    8 => col < 6 ? value : null,
                    4 or 5 or 6 => (row + col) % 3 == 0 ? value : DBNull.Value,
                    10 => col > 4 && col < 10 ? value : "",
                    _ => value
                };
            });

            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            processor.Save();

            var values = InvokeWorkbookScopeWith(filePath, new ReadRow
            {
                SheetName = "Sheet1",
                StartingCell = startingCell,
                Limit = limit
            });
            return values;
        }
    }
}