using Autossential.Workbook.Activities.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class InsertSheetTests : BaseTests
    {
        [Test]
        public void InsertSheet_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new InsertSheet
                {
                    SheetName = ""
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task InsertSheet_AddsSheet(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new InsertSheet
            {
                SheetName = "NewSheet",
                Position = new InArgument<int?>(ctx => 0)
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var names = p2.GetSheetNames();
            await Assert.That(names).Contains("NewSheet");
        }
    }
}