using Autossential.Workbook.Activities.Core;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class AppendRangeTests : BaseTests
    {
        [Test]
        public void AppendRange_Fails_WhenMissingSheet()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new AppendRange
                {
                    SheetName = "",
                    DataTable = new InArgument<DataTable>(ctx => new DataTable())
                });
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task AppendRange_AppendsRows_ToExistingSheet(string extension)
        {
            var initial = TableUtils.Generate(3, 2, 1);
            var toAppend = TableUtils.Generate(3, 2, 2);

            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", initial, "A1", true);
            processor.Save();

            InvokeWorkbookScopeWith(filePath, new AppendRange
            {
                SheetName = "Sheet1",
                DataTable = new InArgument<DataTable>(ctx => toAppend)
            });

            var p2 = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var read = p2.ReadRange("Sheet1", "A1", true);
            await Assert.That(read.Rows.Count).IsEqualTo(4);
        }
    }
}