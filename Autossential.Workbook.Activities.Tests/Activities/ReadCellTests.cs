namespace Autossential.Workbook.Activities.Tests.Activities
{
    //public class ReadCellTests
    //{
    //    [Theory]
    //    [InlineData(true)]
    //    [InlineData(false)]
    //    public void ReadCell_ValidAddress_ShouldReturnValue(bool openXmlFormat)
    //    {
    //        var path = WorkbookGenerator.CreateWorkbookFile(openXmlFormat, sheet =>
    //        {
    //            var header = sheet.CreateRow(0);
    //            header.CreateCell(0).SetCellValue("Text");
    //            header.CreateCell(1).SetCellValue("Number");
    //            var row = sheet.CreateRow(1);
    //            row.CreateCell(0).SetCellValue("Brazil");
    //            row.CreateCell(1).SetCellValue(55);
    //        });

    //        var A1 = new Variable<object>();
    //        var B1 = new Variable<object>();
    //        var A2 = new Variable<object>();
    //        var B2 = new Variable<object>();

    //        var result = WorkbookFixture.InvokeWorkbookScopeWith(path, [A1, B1, A2, B2], [
    //            new ReadCell { SheetName = "Sheet1", CellAddress = "A1", Result = new OutArgument<object>(A1) },
    //            new ReadCell { SheetName = "Sheet1", CellAddress = "B1", Result = new OutArgument<object>(B1) },
    //            new ReadCell { SheetName = "Sheet1", CellAddress = "A2", Result = new OutArgument<object>(A2) },
    //            new ReadCell { SheetName = "Sheet1", CellAddress = "B2", Result = new OutArgument<object>(B2) }],
    //            env => new Dictionary<string, object>
    //            {
    //                { "A1", A1.Get(env) },
    //                { "B1", B1.Get(env) },
    //                { "A2", A2.Get(env) },
    //                { "B2", B2.Get(env) }
    //            });

    //        Assert.Equal("Text", result["A1"]);
    //        Assert.Equal("Number", result["B1"]);
    //        Assert.Equal("Brazil", result["A2"]);
    //        Assert.Equal(55d, result["B2"]);
    //    }

    //    [Theory]
    //    [InlineData(true)]
    //    [InlineData(false)]
    //    public void ReadCell_EmptyCell_ShouldReturnNull(bool openXmlFormat)
    //    {
    //        var path = WorkbookGenerator.CreateWorkbookFile(openXmlFormat, _ => { });

    //        var result = WorkbookFixture.InvokeWorkbookScopeWith(path, new ReadCell
    //        {
    //            SheetName = "Sheet1",
    //            CellAddress = "B1"
    //        });

    //        Assert.Null(result);
    //    }
    //}
}
