using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class WorkbookScopeTests : BaseTests
    {
        [Test]
        [Arguments(WorkbookScope.TAG)]
        [Arguments("_wbInst")]
        public async Task WorkbookScope_ShouldWork_WhenDifferentNameTag(string tag)
        {
            var (processor, filePath) = NewFile(".xlsx");
            processor.WriteCell("Sheet1", "A1", 123d);
            processor.Save();

            var result = InvokeWorkbookScopeWith<object>(filePath, new ReadCell
            {
                SheetName = "Sheet1",
                CellAddress = "A1",
            });

            await Assert.That(result).IsEqualTo(123d);
        }
    }
}