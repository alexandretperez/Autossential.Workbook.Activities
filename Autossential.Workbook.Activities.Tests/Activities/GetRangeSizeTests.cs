namespace Autossential.Workbook.Activities.Tests.Activities
{
    //public class GetRangeSizeTests(WorkbookFixture fixture) : IClassFixture<WorkbookFixture>
    //{
    //    public WorkbookFixture Fixture { get; } = fixture;

    //    [Theory]
    //    [InlineData(true, "A1", 10, 10)]
    //    [InlineData(true, "D5:F12", 6, 3)]
    //    [InlineData(true, "B11:G15", 0, 0)]
    //    [InlineData(false, "A1", 10, 10)]
    //    [InlineData(false, "D6:F12", 5, 3)]
    //    [InlineData(false, "J10:K14", 1, 1)]

    //    public void GetRangeSize_FullFill_ReturnsExpectedDimensions(bool openXmlFormat, string range, int rows, int cols)
    //    {
    //        var path = openXmlFormat ? Fixture.OpenXMLFilePath : Fixture.BinaryFilePath;

    //        var rowCount = new Variable<int>();
    //        var colCount = new Variable<int>();

    //        var result = WorkbookFixture.InvokeWorkbookScopeWith(path, [rowCount, colCount], [new GetRangeSize
    //        {
    //            SheetName = "Sheet1",
    //            Range = range,
    //            RowCount = new OutArgument<int>(rowCount),
    //            ColumnCount = new OutArgument<int>(colCount)
    //        }], env => Tuple.Create(rowCount.Get(env), colCount.Get(env)));

    //        Assert.Equal(rows, result.Item1);
    //        Assert.Equal(cols, result.Item2);
    //    }


    //    [Theory]
    //    [InlineData(true, "Sheet2", "A1:E3", 3, 3)]
    //    [InlineData(false, "Sheet2", "C2:F4", 2, 2)]
    //    [InlineData(true, "Sheet3", "A1:E3", 2, 3)]
    //    [InlineData(false, "Sheet3", "B2:G4", 3, 5)]
    //    public void GetRangeSize_PartialFill_ReturnsExpectedDimensions(bool openXmlFormat, string sheetName, string range, int rows, int cols)
    //    {
    //        var path = openXmlFormat ? Fixture.OpenXMLFilePath : Fixture.BinaryFilePath;

    //        var rowCount = new Variable<int>();
    //        var colCount = new Variable<int>();

    //        var result = WorkbookFixture.InvokeWorkbookScopeWith(path, [rowCount, colCount], [new GetRangeSize
    //        {
    //            SheetName = sheetName,
    //            Range = range,
    //            RowCount = new OutArgument<int>(rowCount),
    //            ColumnCount = new OutArgument<int>(colCount)
    //        }],
    //        env => Tuple.Create(rowCount.Get(env), colCount.Get(env)));

    //        Assert.Equal(rows, result.Item1);
    //        Assert.Equal(cols, result.Item2);
    //    }
    //}
}