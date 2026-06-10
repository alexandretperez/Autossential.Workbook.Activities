namespace Autossential.Workbook.Activities.Tests.Activities
{
    //public class ReadRangeTests : IClassFixture<WorkbookFixture>
    //{
    //    public ReadRangeTests(WorkbookFixture fixture)
    //    {
    //        Fixture = fixture;
    //    }

    //    public WorkbookFixture Fixture { get; }

    //    [Theory]
    //    [InlineData(true, "Sheet1", false, "A1", 10, 10, 1, 1)]
    //    [InlineData(false, "Sheet1", false, "A1:E5", 5, 5, 1, 1)]
    //    [InlineData(true, "Sheet1", true, "A1", 9, 10, 1, 1)]
    //    [InlineData(false, "Sheet4", true, "A3", 9, 5, 2, 3)]
    //    [InlineData(false, "Sheet4", true, "A1", 10, 5, 2, 3)]

    //    public void ReadRange_ReturnsTable(bool openXmlFormat, string sheetName, bool hasHeaders, string range, int rows, int cols, int headerRows, int rowsPerRecord)
    //    {
    //        var path = openXmlFormat ? Fixture.OpenXMLFilePath : Fixture.BinaryFilePath;
    //        var result = WorkbookFixture.InvokeWorkbookScopeWith(path, new ReadRange
    //        {
    //            SheetName = sheetName,
    //            HasHeaders = hasHeaders,
    //            Range = range,
    //            RowsPerRecord = rowsPerRecord,
    //            HeaderRows = headerRows
    //        });

    //        Assert.Equal(rows, result.Rows.Count);
    //        Assert.Equal(cols, result.Columns.Count);
    //    }
    //}
}
