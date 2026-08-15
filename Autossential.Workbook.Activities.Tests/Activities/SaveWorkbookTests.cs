using Autossential.Workbook.Activities.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class SaveWorkbookTests : BaseTests
    {
        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task SaveWorkbook_PersistsChanges(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.Save();

            var write = new WriteCell
            {
                SheetName = "Sheet1",
                CellAddress = "B2",
                Value = new InArgument<object>(ctx => "SavedValue")
            };

            InvokeWorkbookScopeWith(filePath, [], [write, new SaveWorkbook()], env => "ok");

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var value = p2.ReadCell("Sheet1", "B2");
            await Assert.That(value?.ToString()).IsEqualTo("SavedValue");
        }
    }
}