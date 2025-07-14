using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Activities;
using System.Activities.Statements;

namespace Autossential.Workbook.Tests.Activities
{
    [TestClass]
    public class GetSheetNames_Tests
    {
        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetSheetNames_FromFile_ReturnsCorrectSheetNames(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var dyn = new DynamicActivity<string[]>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new GetSheetNames
                    {
                        Result = new OutArgument<string[]>(env => dyn.Result.Get(env))
                    }
                }
            };

            CollectionAssert.AreEqual(new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" }, WorkflowInvoker.Invoke(dyn));
        }


        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetSheetNames_FromMemory_ReturnsCorrectSheetNames(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var dyn = new DynamicActivity<string[]>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                AutoSave = false,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new Sequence
                    {
                        Activities =
                        {
                            new RenameSheet
                            {
                                SheetNameOrIndex = new InArgument<int>(0),
                                NewSheetName = "New"
                            },
                            new GetSheetNames
                            {
                                Result = new OutArgument<string[]>(env => dyn.Result.Get(env))
                            }
                        }
                    }
                }
            };

            CollectionAssert.AreEqual(new[] { "New", "Sheet2", "Sheet3", "Sheet4" }, WorkflowInvoker.Invoke(dyn));
        }
    }
}