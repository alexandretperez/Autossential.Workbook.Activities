using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class DeleteRowsTests : BaseTests
    {
        [Test]
        public void DeleteRows_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new DeleteRows
                {
                    SheetName = "",
                    RowReferences = "1"
                });
            });
        }

        [Test]
        public void DeleteRows_Fails_WhenRowReferencesIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new DeleteRows
                {
                    SheetName = "Sheet1",
                    RowReferences = ""
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task DeleteRows_RemovesExpectedRows(string extension)
        {
            var data = TableUtils.Build(5, 5, (c, r) => $"C{c}R{r}");
            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            processor.Save();

            // Remove rows 1 and range 3:4 => removes 3 rows
            InvokeWorkbookScopeWith(filePath, new DeleteRows
            {
                SheetName = "Sheet1",
                RowReferences = "1,3:4"
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var rows = p2.GetRowCount("Sheet1", "A1");
            await Assert.That(rows).IsEqualTo(2);
        }
    }
}