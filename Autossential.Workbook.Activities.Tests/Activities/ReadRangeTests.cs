using Autossential.Workbook.Activities.Core;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;
using System.Activities.Statements;
using System.Data;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadRangeTests : IClassFixture<WorkbookFixture>
    {
        public ReadRangeTests(WorkbookFixture fixture)
        {
            Fixture = fixture;
        }

        public WorkbookFixture Fixture { get; }

        [Theory]
        //[InlineData(true, "Sheet1", false, "A1", 10, 10, 1, 1)]
        //[InlineData(false, "Sheet1", false, "A1:E5", 5, 5, 1, 1)]
        //[InlineData(true, "Sheet1", true, "A1", 9, 10, 1, 1)]
        [InlineData(false, "Sheet4", true, "A3", 9, 5, 2, 3)]
        //[InlineData(false, "Sheet4", true, "A1", 10, 5, 2, 3)]

        public void ReadRange_ReturnsTable(bool openXmlFormat, string sheetName, bool hasHeaders, string range, int rows, int cols, int headerRows, int rowsPerRecord)
        {
            var path = openXmlFormat ? Fixture.OpenXMLFilePath : Fixture.BinaryFilePath;

            var table = new Variable<DataTable>();

            var dyn = new DynamicActivity<DataTable>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new Sequence
                    {
                        Variables = { table },
                        Activities =
                        {
                            new ReadRange
                            {
                                SheetName= sheetName,
                                HasHeaders = hasHeaders,
                                Range = range,
                                RowsPerRecord = rowsPerRecord,
                                HeaderRows = headerRows,
                                Result = new OutArgument<DataTable>(env=>dyn.Result.Get(env))
                            },
                        }
                    }
                }
            };

            var result = WorkflowInvoker.Invoke(dyn);
            Assert.Equal(rows, result.Rows.Count);
            Assert.Equal(cols, result.Columns.Count);
            if (hasHeaders)
            {
                Assert.NotEqual("Col1", result.Columns[0].ColumnName);
            }
            else
            {
                Assert.Equal("Col1", result.Columns[0].ColumnName);
            }
        }
    }
}
