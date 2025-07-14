using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Activities;

namespace Autossential.Workbook.Tests.Activities
{
    [TestClass]
    public class GetRowCountTests
    {
        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", 11)]
        [DataRow("OXML_data.xlsx", "A2", 10)]

        [DataRow("BIFF8_data.xls", "A1", 11)]
        [DataRow("BIFF8_data.xls", "A2", 10)]
        public void GetRowCount(string fileName, string range, int expectedCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var dyn = new DynamicActivity<int>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new GetRowCount
                    {
                        SheetName = "Header",
                        Range = range,
                        Result = new OutArgument<int>(env => dyn.Result.Get(env))
                    }
                }
            };

            var result = WorkflowInvoker.Invoke(dyn);
            Assert.AreEqual(expectedCount, result);
        }
    }
}
