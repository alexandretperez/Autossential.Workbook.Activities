using Autossential.Workbook.Activities.Core;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;
using System.Activities.Statements;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadCellTests
    {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadCell_ValidAddress_ShouldReturnValue(bool openXmlFormat)
        {
            var path = WorkbookGenerator.CreateWorkbookFile(openXmlFormat, sheet =>
            {
                var header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("Text");
                header.CreateCell(1).SetCellValue("Number");
                var row = sheet.CreateRow(1);
                row.CreateCell(0).SetCellValue("Brazil");
                row.CreateCell(1).SetCellValue(55);
            });

            var A1 = new Variable<object>();
            var B1 = new Variable<object>();
            var A2 = new Variable<object>();
            var B2 = new Variable<object>();

            var dyn = new DynamicActivity<IDictionary<string, object>>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new Sequence
                    {
                        Variables = { A1, B1, A2, B2 },
                        Activities =
                        {
                            new ReadCell { SheetName = "Sheet1", CellAddress = "A1", Result = new OutArgument<object>(A1) },
                            new ReadCell { SheetName = "Sheet1", CellAddress = "B1", Result = new OutArgument<object>(B1) },
                            new ReadCell { SheetName = "Sheet1", CellAddress = "A2", Result = new OutArgument<object>(A2) },
                            new ReadCell { SheetName = "Sheet1", CellAddress = "B2", Result = new OutArgument<object>(B2) },
                            new Assign<IDictionary<string, object>>
                            {
                                To   = new OutArgument<IDictionary<string, object>>(env => dyn.Result.Get(env)),
                                Value = new InArgument<IDictionary<string, object>>(env => new Dictionary<string, object>
                                {
                                    { "A1", A1.Get(env) },
                                    { "B1", B1.Get(env) },
                                    { "A2", A2.Get(env) },
                                    { "B2", B2.Get(env) }
                                })
                            }
                        }
                    }
                }
            };

            var result = WorkflowInvoker.Invoke(dyn);

            Assert.Equal("Text", result["A1"]);
            Assert.Equal("Number", result["B1"]);
            Assert.Equal("Brazil", result["A2"]);
            Assert.Equal(55d, result["B2"]);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadCell_EmptyCell_ShouldReturnNull(bool openXmlFormat)
        {
            var path = WorkbookGenerator.CreateWorkbookFile(openXmlFormat, _ => { });

            var dyn = new DynamicActivity<object>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new Sequence
                    {
                        Activities =
                        {
                            new ReadCell {
                                SheetName = "Sheet1",
                                CellAddress = "B1",
                                Result = new OutArgument<object>(env => dyn.Result.Get(env) )
                            }
                        }
                    }
                }
            };

            var result = WorkflowInvoker.Invoke<object>(dyn);
            Assert.Null(result);
        }
    }
}
