using Autossential.Workbook.Activities.Core;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class WorkbookProcessorTests
    {
        [Theory]
        [InlineData("", 7, 5)]
        [InlineData("A1", 7, 5)]
        [InlineData("A1:G4", 3, 7)]
        [InlineData("B1:F20", 19, 5)]
        [InlineData("C3:F6", 3, 4)]
        public void ReadRange(string range, int rows, int cols)
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox.xlsx");
            var table = processor.ReadRange("Planilha1", range, true, 1, 1);
            Assert.Equal(rows, table.Rows.Count);
            Assert.Equal(cols, table.Columns.Count);
        }
    }
}
