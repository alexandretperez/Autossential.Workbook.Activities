using Autossential.Workbook.Activities.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class WriteCellTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "A1", "Hello")]
        [Arguments(".xls", "B2", 1)]

        [Arguments(".xlsx", "A1", "Hello")]
        [Arguments(".xlsx", "B2", 1)]
        public async Task WriteCell_CellIsUpdated_BasedOnArguments(string extension, string cell, object value)
        {
            object readValue = Run(extension, cell, value);

            await Assert.That(value?.ToString()).IsEqualTo(readValue?.ToString());
        }

        [Test]

        [Arguments(".xls", null)]
        [Arguments(".xls", "")]

        [Arguments(".xlsx", null)]
        [Arguments(".xlsx", "")]
        public void WriteCell_Fails_WhenMissingCell(string extension, string? cell)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() => Run(extension, cell, 0));
        }

        [Test]
        public void WriteCell_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new WriteCell
                {
                    SheetName = "",
                    Value = new InArgument<object>(ctx => 1),
                    CellAddress = ""
                });
            });
        }

        private object Run(string extension, string? cell, object value)
        {
            var filePath = NewTempFilePath(extension);
            InvokeWorkbookScopeWith(filePath, new WriteCell
            {
                SheetName = "Sheet1",
                CellAddress = cell,
                Value = new InArgument<object>(ctx => value)
            });

            var processor = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var readValue = processor.ReadCell("Sheet1", cell);
            return readValue;
        }
    }
}