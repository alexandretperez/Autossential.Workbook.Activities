using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadCellTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "")]
        [Arguments(".xls", null)]

        [Arguments(".xlsx", "")]
        [Arguments(".xlsx", null)]
        public void ReadCell_Fails_WhenCellAddressIsMissing(string extension, string? cell)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() => Run(extension, cell, "A1"));
        }

        [Test]
        public void ReadCell_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith<object>(NewTempFilePath(".xlsx"), new ReadCell
                {
                    SheetName = "",
                    CellAddress = ""
                });
            });
        }

        [Test]

        [Arguments(".xls", "B3")]
        [Arguments(".xls", "J14")]
        [Arguments(".xls", "A1")]

        [Arguments(".xlsx", "B3")]
        [Arguments(".xlsx", "J14")]
        [Arguments(".xlsx", "A1")]
        public async Task ReadCell_ReturnsExpectedValue_AfterWrite(string extension, string cell)
        {
            Dictionary<string, object> result = Run(extension, cell, cell);

            await Assert.That(result["Value"].ToString()).IsEqualTo("Hello");
        }

        private Dictionary<string, object> Run(string extension, string? cell, string? writeCell)
        {
            var (processor, filePath) = NewFile(extension);
            processor.WriteCell("Sheet1", writeCell, "Hello");
            processor.Save();

            var value = new Variable<object>();
            var result = InvokeWorkbookScopeWith(filePath, [value], [new ReadCell {
                SheetName = "Sheet1",
                CellAddress = cell,
                Result = new OutArgument<object>(value)
            }], env => new Dictionary<string, object>
            {
                {"Value",value.Get(env) }
            });
            return result;
        }
    }
}