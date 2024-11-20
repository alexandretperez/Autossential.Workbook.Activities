using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Data;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class ReadRangeTest
    {
        [TestMethod]
   //     [DataRow("openxml.xlsx")]
        [DataRow("ole2.xls")]
        public void Read(string fileName)
        {
            var readRange = new ReadRange()
            {
                AddHeaders = true,
                UseScope = false,
            };

            var result = WorkflowTester.Run(readRange, CreateArgs(IOSamples.GetSamplePath(fileName)));
            var dt = (DataTable)result.Get(p => p.Result);
            Assert.AreEqual(dt.Rows.Count, 10);
            Assert.AreEqual(dt.Columns.Count, 6);
        }

        private static Dictionary<string, object> CreateArgs(string filePath)
        {
            return new Dictionary<string, object>
            {
                { nameof(ReadRange.WorkbookPath), filePath },
                { nameof(ReadRange.SheetName), "Sheet1" },
                { nameof(ReadRange.Range), "" }
            };
        }
    }
}