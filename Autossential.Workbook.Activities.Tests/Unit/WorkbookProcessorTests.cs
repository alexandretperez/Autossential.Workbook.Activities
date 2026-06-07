using Autossential.Workbook.Activities.Core;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class WorkbookProcessorTests
    {
        private string filePath = @"C:\Users\alexa\Downloads\Sandbox_copy.xlsx";
        public WorkbookProcessorTests()
        {
            if (!File.Exists(filePath))
                File.Copy(@"C:\Users\alexa\Downloads\Sandbox.xlsx", filePath);
        }

        [Theory]
        [InlineData("Empty", "B3:E9", 7, 4, false)]
        [InlineData("Empty", "B3:E9", 6, 4, true)]
        [InlineData("Empty", "B3", 0, 0, false)]
        [InlineData("Empty", "", 0, 0, true)]

        [InlineData("Companies", "", 11, 7, false)]
        [InlineData("Companies", "A1", 10, 7, true)]
        [InlineData("Companies", "C1", 10, 5, true)]

        [InlineData("Animals", "A1", 8, 4, true)]
        [InlineData("Animals", "B4", 5, 3, true)]

        [InlineData("Cars", "A1", 8, 4, false)]
        [InlineData("Cars", "B3:F11", 8, 5, true)]
        [InlineData("Cars", "A1:E10", 10, 5, false)]
        public void ReadRange(string sheetName, string range, int rows, int cols, bool hasHeaders)
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox_copy.xlsx");
            var table = processor.ReadRange(sheetName, range, hasHeaders, 1, 1);
            Assert.Equal(rows, table.Rows.Count);
            Assert.Equal(cols, table.Columns.Count);
        }

        [Theory]
        [InlineData("Bug", "A1", 0, 7)]
        [InlineData("Bug", "A2", 0, 6)]
        [InlineData("Bug", "A3", 0, 5)]
        [InlineData("Bug", "A4", 0, 7)]

        [InlineData("Bug", "B1", 0, 6)]
        [InlineData("Bug", "C2", 0, 4)]
        [InlineData("Bug", "D3", 0, 2)]
        [InlineData("Bug", "E4", 0, 3)]

        [InlineData("Bug", "A1", 4, 4)]
        [InlineData("Bug", "B1", 1, 1)]
        [InlineData("Bug", "B2", 1, 0)]
        [InlineData("Bug", "C3", 2, 2)]

        [InlineData("Companies", "A1", 0, 7)]
        public void ReadRow(string sheetName, string startingCell, int limit, int expectedCount)
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox_copy.xlsx");
            var values = processor.ReadRow(sheetName, startingCell, limit);
            Assert.Equal(expectedCount, values.Length);
        }

        [Theory]
        [InlineData("Bug", "A1", 0, 2)]
        [InlineData("Bug", "B1", 0, 3)]
        [InlineData("Bug", "C1", 0, 4)]
        [InlineData("Bug", "D1", 0, 3)]

        [InlineData("Bug", "A3", 0, 0)]
        [InlineData("Bug", "B3", 0, 1)]
        [InlineData("Bug", "D2", 0, 2)]
        [InlineData("Bug", "G2", 0, 3)]

        [InlineData("Bug", "B2", 2, 2)]
        [InlineData("Bug", "E2", 1, 0)]
        [InlineData("Bug", "G1", 3, 1)]

        [InlineData("Companies", "A1", 0, 11)]
        public void ReadColumn(string sheetName, string startingCell, int limit, int expectedCount)
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox_copy.xlsx");
            var values = processor.ReadColumn(sheetName, startingCell, limit);
            Assert.Equal(expectedCount, values.Length);
        }

        [Fact]
        public void WriteRange()
        {
            var dt = TableGenerator.GenerateTable(5, 5, typeof(string), typeof(int), typeof(DateTime), typeof(bool), typeof(double));
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox_copy.xlsx");
            processor.WriteRange("Bug", dt, "B2", true);
            processor.Save();
            processor.Dispose();
        }
    }
}
