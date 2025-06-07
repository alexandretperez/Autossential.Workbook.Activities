using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{

    [TestClass]
    public class GetRowCount_Tests
    {
        [TestMethod]
        [DataRow("OXML_scatter.xlsx", "A1", 6)]
        [DataRow("OXML_scatter.xlsx", "A1:E5", 3)]
        [DataRow("BIFF8_scatter.xls", "B2:D5", 2)]
        [DataRow("OXML_scatter.xlsx", "D3:E6", 4)]

        [DataRow("BIFF8_scatter.xls", "A1", 6)]
        [DataRow("BIFF8_scatter.xls", "A1:E5", 3)]
        [DataRow("BIFF8_scatter.xls", "B2:D5", 2)]
        [DataRow("BIFF8_scatter.xls", "D3:E6", 4)]
        public void GetRowCount_SheetNameExists_ReturnRowCount(string fileName, string range, int expectedCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var count = workbook.GetRowCount("Sheet1", range);
            workbook.Dispose();
            Assert.AreEqual(expectedCount, count);
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetRowCount_InvalidSheetName_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.GetRowCount("", "A1"));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetRowCount_InvalidRange_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.GetRowCount("Sheet1", "A"));
            workbook.Dispose();
        }
    }
}
