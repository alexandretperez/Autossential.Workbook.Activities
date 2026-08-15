using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class RenameSheetTests : BaseTests
    {
        [Test]
        public void RenameSheet_Fails_WhenSheetIsMissingOrNewNameMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new RenameSheet
                {
                    SheetName = "",
                    NewSheetName = "X"
                });
            });

            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new RenameSheet
                {
                    SheetName = "A",
                    NewSheetName = ""
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task RenameSheet_RenamesSheet(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.InsertSheet("OldName");
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new RenameSheet
            {
                SheetName = "OldName",
                NewSheetName = "Renamed"
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var names = p2.GetSheetNames();
            await Assert.That(names).Contains("Renamed");
        }
    }
}