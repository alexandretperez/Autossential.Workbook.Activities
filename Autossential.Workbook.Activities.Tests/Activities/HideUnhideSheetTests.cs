using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class HideUnhideSheetTests : BaseTests
    {
        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task HideUnhideSheet_HideAndUnhide_NoException(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new Workbook.Activities.HideUnhideSheet
            {
                SheetName = "Sheet1",
                Action = Workbook.Activities.HideUnhideSheet.HideUnhideAction.Hide
            });

            InvokeWorkbookScopeWith(filePath, new Workbook.Activities.HideUnhideSheet
            {
                SheetName = "Sheet1",
                Action = Workbook.Activities.HideUnhideSheet.HideUnhideAction.Unhide
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var names = p2.GetSheetNames();
            await Assert.That(names).Contains("Sheet1");
        }



        [Test]
        public void HideUnhideSheet_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new Workbook.Activities.HideUnhideSheet
                {
                    SheetName = "",
                    Action = Workbook.Activities.HideUnhideSheet.HideUnhideAction.Hide
                });
            });
        }
    }
}