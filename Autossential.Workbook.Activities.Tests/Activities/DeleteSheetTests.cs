using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class DeleteSheetTests : BaseTests
    {
        [Test]
        public void DeleteSheet_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new DeleteSheet
                {
                    SheetName = ""
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task DeleteSheet_RemovesSheet(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.InsertSheet("ToDelete");
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new DeleteSheet
            {
                SheetName = "ToDelete"
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var names = p2.GetSheetNames();
            await Assert.That(names).DoesNotContain("ToDelete");
        }
    }
}