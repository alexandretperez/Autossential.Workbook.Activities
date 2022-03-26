using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class GetSheetNamesTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();
            IOSamples.ExportSample("book.xls");
            IOSamples.ExportSample("book.xlsx");
        }

        [ClassCleanup]
        public static void CleanUp()
        {
            IOSamples.ClearFolder();
        }

        [TestMethod("GetSheetNames")]
        [DataRow("book.xlsx")]
        public void OpenXML(string file)
        {
            var path = IOSamples.GetTestPath(file);
            var args = new Dictionary<string, object>
            {
                { nameof(GetSheetNames.WorkbookPath), path }
            };

            var result = WorkflowTester.Run(new GetSheetNames(), args);
            var values = result.Get(p => p.Result) as string[];
            CollectionAssert.AreEqual(new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" }, values);
        }

        [TestMethod("GetSheetNames")]
        [DataRow("book.xls")]
        public void OLE2(string file)
        {
            var path = IOSamples.GetTestPath(file);
            var args = new Dictionary<string, object>
            {
                { nameof(GetSheetNames.WorkbookPath), path }
            };

            var result = WorkflowTester.Run(new GetSheetNames(), args);
            var values = result.Get(p => p.Result) as string[];
            CollectionAssert.AreEqual(new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" }, values);
        }
    }
}
