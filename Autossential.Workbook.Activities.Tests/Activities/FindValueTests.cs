using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class FindValueTests : BaseTests
    {
        [Test]

        [Arguments(".xls", "A1", "Col3", "C1", 3, 1)]
        [Arguments(".xls", "B2", 657.52, "B2", 2, 2)]
        [Arguments(".xls", "A1", "IamNotThere", "", -1, -1)]

        [Arguments(".xlsx", "A1", "Col3", "C1", 3, 1)]
        [Arguments(".xlsx", "B2", 657.52, "B2", 2, 2)]
        [Arguments(".xlsx", "A1", "IamNotThere", "", -1, -1)]
        public async Task FindValue_ReturnsCoordinates_WhenValidArguments(string extension, string range, object? value, string expectedAddress, int expectedCol, int expectedRow)
        {
            Tuple<string, int, int> result = Run(extension, range, value);

            await Assert.That(result.Item1).IsEqualTo(expectedAddress);
            await Assert.That(result.Item2).IsEqualTo(expectedCol);
            await Assert.That(result.Item3).IsEqualTo(expectedRow);
        }

        [Test]

        [Arguments(".xls", "", "Col1", "A1", 1, 1)]
        [Arguments(".xls", null, "Col1", "A1", 1, 1)]

        [Arguments(".xlsx", "", "Col1", "A1", 1, 1)]
        [Arguments(".xlsx", null, "Col1", "A1", 1, 1)]
        public async Task FindValue_ReturnsCoordinatesBaseInWholeSheet_WhenMissingRange(string extension, string? range, object? value, string expectedAddress, int expectedCol, int expectedRow)
        {
            Tuple<string, int, int> result = Run(extension, range, value);

            await Assert.That(result.Item1).IsEqualTo(expectedAddress);
            await Assert.That(result.Item2).IsEqualTo(expectedCol);
            await Assert.That(result.Item3).IsEqualTo(expectedRow);
        }

        [Test]
        public void FindValue_Fails_WhenSheetIsMissing()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                InvokeWorkbookScopeWith(NewTempFilePath(".xlsx"), new FindValue
                {
                    SheetName = "",
                    Value = new InArgument<object>(ctx => 1)
                });
            });
        }

        private Tuple<string, int, int> Run(string extension, string? range, object? value)
        {
            var data = TableUtils.Generate(5, 5, 1);
            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", true);
            processor.Save();

            var address = new Variable<string>();
            var col = new Variable<int>();
            var row = new Variable<int>();

            var result = InvokeWorkbookScopeWith(filePath, [address, col, row], [
                new FindValue
                {
                    SheetName="Sheet1",
                    Range = range,
                    Value = new InArgument<object>(ctx => value!),
                    CellAddress = new OutArgument<string>(address),
                    ColumnNumber = new OutArgument<int>(col),
                    RowNumber = new OutArgument<int>(row)
                }
             ], env => Tuple.Create(address.Get(env), col.Get(env), row.Get(env)));
            return result;
        }
    }
}