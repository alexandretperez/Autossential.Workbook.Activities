using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{

    [TestClass]
    public class GetColumnCount_Tests
    {
        [TestMethod]
        [DataRow("OXML_scatter.xlsx", "A1", 7)]
        [DataRow("OXML_scatter.xlsx", "A1:E5", 4)]
        [DataRow("OXML_scatter.xlsx", "A1:B5", 1)]

        [DataRow("BIFF8_scatter.xls", "A1", 7)]
        [DataRow("BIFF8_scatter.xls", "A1:E5", 4)]
        [DataRow("BIFF8_scatter.xls", "A1:B5", 1)]
        public void GetColumnCount_SheetNameExists_ReturnColumnCount(string fileName, string range, int expectedCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var columnsCount = workbook.GetColumnCount("Sheet1", range);
            workbook.Dispose();
            Assert.AreEqual(expectedCount, columnsCount);
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetColumnCount_InvalidSheetName_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.GetColumnCount("", "A1"));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetColumnCount_InvalidRange_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.GetColumnCount("Sheet1", "A"));
            workbook.Dispose();
        }
    }
}
