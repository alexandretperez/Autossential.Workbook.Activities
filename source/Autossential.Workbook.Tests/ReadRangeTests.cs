using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HPSF;
using System;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class ReadRange_Tests
    {

        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", true, 10, 6)]
        //[DataRow("OXML_data.xlsx", "A1", false, 11, 6)]
        [DataRow("OXML_data.xlsx", "A1:B5", true, 4, 2)]

        [DataRow("BIFF8_data.xls", "A1", true, 10, 6)]
        //[DataRow("BIFF8_data.xls", "A1", false, 11, 6)]
        [DataRow("BIFF8_data.xls", "A1:B5", true, 4, 2)]
        public void ReadRange(string fileName, string range, bool hasHeaders, int rowCount, int colCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var dyn = new DynamicActivity<DataTable>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new ReadRange
                    {
                        SheetName = "Header",
                        Range = range,
                        HasHeaders = hasHeaders,
                        Result = new OutArgument<DataTable>(env => dyn.Result.Get(env))
                    }
                }
            };

            var dt = WorkflowInvoker.Invoke(dyn);

            Assert.IsNotNull(dt);
            Assert.AreEqual(rowCount, dt.Rows.Count);
            Assert.AreEqual(colCount, dt.Columns.Count);
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", 10, 6)]
     
        public void ReadRange_NoHeader(string fileName, string range, int rowCount, int colCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var dyn = new DynamicActivity<DataTable>();
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = path,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                    Handler = new ReadRange
                    {
                        SheetName = "NoHeader",
                        Range = range,
                        HasHeaders = false,
                        Result = new OutArgument<DataTable>(env => dyn.Result.Get(env))
                    }
                }
            };

            var dt = WorkflowInvoker.Invoke(dyn);

            Assert.IsNotNull(dt);
            Assert.AreEqual(rowCount, dt.Rows.Count);
            Assert.AreEqual(colCount, dt.Columns.Count);
        }
    }
}
