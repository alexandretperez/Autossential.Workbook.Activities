namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class FreezePanesTests : BaseTests
    {
        [Test]
        public void FreezePanes_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new FreezePanes
                {
                    SheetName = "",
                    ColumnsToFreeze = 1,
                    RowsToFreeze = 1
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public void FreezePanes_DoesNotThrow_ForValidSheet(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new FreezePanes
            {
                SheetName = "Sheet1",
                ColumnsToFreeze = -1,
                RowsToFreeze = -5
            });
        }
    }
}