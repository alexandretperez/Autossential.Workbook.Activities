using Autossential.Workbook.Activities.Core;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class WriteRangeTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "A1", true, 5, 5)]
        [Arguments(".xls", "C14", true, 5, 5)]

        [Arguments(".xlsx", "A1", true, 5, 5)]
        [Arguments(".xlsx", "C14", true, 5, 5)]
        public async Task WriteRange_ExpectedRange_BasedOnArguments(string extension, string? startingCell, bool addHeaders, int expectedCols, int expectedRows)
        {
            var readData = Run(extension, startingCell, addHeaders);

            await Assert.That(readData.Rows.Count).IsEqualTo(expectedRows);
            await Assert.That(readData.Columns.Count).IsEqualTo(expectedCols);
        }

        [Test]
        public void WriteRange_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new WriteRange
                {
                    SheetName = "",
                    DataTable = new InArgument<DataTable>(ctx => new DataTable()),
                    StartingCell = ""
                });
            });
        }

        [Test]

        [Arguments(".xls", "", true, 5, 5)]
        [Arguments(".xls", null, true, 5, 5)]

        [Arguments(".xlsx", "", true, 5, 5)]
        [Arguments(".xlsx", null, true, 5, 5)]
        public async Task WriteRange_Fails_WhenMissingStartingCell(string extension, string? startingCell, bool addHeaders, int expectedCols, int expectedRows)
        {
            DataTable readData = Run(extension, startingCell, addHeaders);

            await Assert.That(readData.Rows.Count).IsEqualTo(expectedRows);
            await Assert.That(readData.Columns.Count).IsEqualTo(expectedCols);
        }

        private DataTable Run(string extension, string? startingCell, bool addHeaders)
        {
            var data = TableUtils.Generate(5, 5, 2);

            var filePath = NewTempFilePath(extension);
            InvokeWorkbookScopeWith(filePath, new WriteRange
            {
                SheetName = "Sheet1",
                AddHeaders = addHeaders,
                StartingCell = startingCell,
                DataTable = new InArgument<DataTable>(ctx => data)
            });

            var processor = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var readData = processor.ReadRange("Sheet1", startingCell, addHeaders);
            return readData;
        }
    }
}