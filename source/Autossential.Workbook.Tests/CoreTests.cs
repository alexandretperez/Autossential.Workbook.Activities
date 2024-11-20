using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Autossential.Workbook.Tests
{

    [TestClass]
    public class CoreTests
    {
        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", 11)]
        [DataRow("OXML_data.xlsx", "B5", 7)]
        [DataRow("OXML_scatter.xlsx", "A1", 6)]
        [DataRow("OXML_scatter.xlsx", "A1:E5", 3)]
        [DataRow("OXML_scatter.xlsx", "B2:D5", 2)]

        [DataRow("BIFF8_data.xls", "A1", 11)]
        [DataRow("BIFF8_data.xls", "B5", 7)]
        [DataRow("BIFF8_scatter.xls", "A1", 6)]
        [DataRow("BIFF8_scatter.xls", "A1:E5", 3)]
        [DataRow("BIFF8_scatter.xls", "B2:D5", 2)]
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


        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", 6)]
        [DataRow("OXML_data.xlsx", "B3:D5", 3)]
        [DataRow("OXML_scatter.xlsx", "A1", 7)]
        [DataRow("OXML_scatter.xlsx", "A1:E5", 4)]
        [DataRow("OXML_scatter.xlsx", "A1:B5", 1)]

        [DataRow("BIFF8_data.xls", "A1", 6)]
        [DataRow("BIFF8_data.xls", "B3:D5", 3)]
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

        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", true, 10, 6)]
        [DataRow("OXML_data.xlsx", "A1", false, 11, 6)]
        [DataRow("OXML_data.xlsx", "A1:B5", true, 4, 2)]

        [DataRow("BIFF8_data.xls", "A1", true, 10, 6)]
        [DataRow("BIFF8_data.xls", "A1", false, 11, 6)]
        [DataRow("BIFF8_data.xls", "A1:B5", true, 4, 2)]
        public void ReadRange_SheetNameExists_ReturnRange(string fileName, string range, bool hasHeaders, int expectedRowCount, int expectedColumnCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("Sheet1", range, hasHeaders);
            workbook.Dispose();
            Assert.AreEqual(expectedRowCount, dataTable.Rows.Count);
            Assert.AreEqual(expectedColumnCount, dataTable.Columns.Count);
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_SheetNameDoesNotExist_ThrowException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("DoesNotExist", "A1", false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidSheetName_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("", "A1", false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidRange_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("Sheet1", "A", false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetSheetNames_MultipleSheets_ReturnSheetNames(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var sheetNames = workbook.GetSheetNames();
            workbook.Dispose();
            CollectionAssert.AreEquivalent(sheetNames, new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" });
        }
    }
}
