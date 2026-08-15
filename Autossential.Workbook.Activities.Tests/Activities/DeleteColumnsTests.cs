using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class DeleteColumnsTests : BaseTests
    {
        [Test]
        public void DeleteColumns_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new DeleteColumns
                {
                    SheetName = "",
                    ColumnReferences = "A"
                });
            });
        }

        [Test]
        public void DeleteColumns_Fails_WhenColumnReferencesIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new DeleteColumns
                {
                    SheetName = "Sheet1",
                    ColumnReferences = ""
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task DeleteColumns_RemovesExpectedColumns(string extension)
        {
            var data = TableUtils.Build(5, 5, (c, r) => $"C{c}R{r}");
            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            processor.Save();

            // Remove column B and range D:E => removes 3 columns
            InvokeWorkbookScopeWith(filePath, new DeleteColumns
            {
                SheetName = "Sheet1",
                ColumnReferences = "B,D:E"
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var cols = p2.GetColumnCount("Sheet1", "A1");
            await Assert.That(cols).IsEqualTo(2);
        }
    }
}