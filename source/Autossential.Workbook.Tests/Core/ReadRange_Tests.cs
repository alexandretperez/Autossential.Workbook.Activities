using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Linq;

namespace Autossential.Workbook.Tests.Core
{
    [TestClass]
    public class ReadRange_Tests
    {
        [TestMethod]
        [DataRow("OXML_data.xlsx", "Header", "A1", 10, 6)]
        [DataRow("OXML_data.xlsx", "Header", "B1", 10, 5)]
        [DataRow("BIFF8_data.xls", "Header", "A1", 10, 6)]
        [DataRow("BIFF8_data.xls", "Header", "B1", 10, 5)]

        public void ReadRange_WithHeaderAndUseColumnType_ReturnRange(string fileName, string sheetName, string range, int expectedRowCount, int expectedColumnCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange(sheetName, range, true, true);
            workbook.Dispose();
            Assert.AreEqual(expectedRowCount, dataTable.Rows.Count);
            Assert.AreEqual(expectedColumnCount, dataTable.Columns.Count);

            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.ColumnName).ToArray();
            if (range.EndsWith("1") && range.Length == 2)
            {
                Assert.IsTrue(columnTypes.Any(name => name == "Number2"));
            }
            else
            {
                Assert.IsTrue(columnTypes.All(name => name != "Number2"));
            }
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", "NoHeader", "A1", 10, 6)]
        [DataRow("OXML_data.xlsx", "NoHeader", "B1", 10, 5)]
        [DataRow("OXML_data.xlsx", "NoHeader", "B4", 7, 5)]

        [DataRow("BIFF8_data.xls", "NoHeader", "A1", 10, 6)]
        [DataRow("BIFF8_data.xls", "NoHeader", "B1", 10, 5)]
        [DataRow("BIFF8_data.xls", "NoHeader", "B4", 7, 5)]

        public void ReadRange_WithoutHeaderAndUseColumnType_ReturnRange(string fileName, string sheetName, string range, int expectedRowCount, int expectedColumnCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange(sheetName, range, false, true);
            workbook.Dispose();
            Assert.AreEqual(expectedRowCount, dataTable.Rows.Count);
            Assert.AreEqual(expectedColumnCount, dataTable.Columns.Count);

            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.ColumnName).ToArray();
            Assert.IsTrue(columnTypes.Any(name => name == "Col1"));
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_SheetNameDoesNotExist_ThrowException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("DoesNotExist", "A1", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidSheetName_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("", "A1", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidRange_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("Header", "A", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", false)]
        [DataRow("OXML_data.xlsx", true)]
        [DataRow("BIFF8_data.xls", false)]
        [DataRow("BIFF8_data.xls", true)]
        public void ReadRange_ColumnTypes_ReturnRange(string fileName, bool useColumnTypes)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("Header", "A1", true, useColumnTypes);
            workbook.Dispose();

            Assert.AreEqual(6, dataTable.Columns.Count);
            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.DataType).ToArray();

            if (useColumnTypes)
            {
                CollectionAssert.AreEqual(columnTypes, new[]
                {
                    typeof(double),
                    typeof(double),
                    typeof(double),
                    typeof(DateTime),
                    typeof(string),
                    typeof(bool)
                });
            }
            else
            {
                Assert.IsTrue(columnTypes.All(p => p == typeof(object)));
            }
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", false)]
        public void ReadRange_NoHeader_TestColumns(string fileName, bool useColumnTypes)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("NoHeader", "A1", false, useColumnTypes);
            workbook.Dispose();
            Assert.IsTrue(true);
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx")]
        [DataRow("BIFF8_data.xls")]
        public void ReadRange_BooleanTypes_TestColumns(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("Boolean", "A1", true, true);
            workbook.Dispose();
            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.DataType).ToArray();
            CollectionAssert.AreEqual(columnTypes, new[] { typeof(string), typeof(bool) });
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx")]
        [DataRow("BIFF8_data.xls")]
        public void ReadRange_NotBoolean_Test(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("NotBoolean", "A1", true, true);
            workbook.Dispose();
            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.DataType).ToArray();

            Assert.IsTrue(columnTypes.All(type => type == typeof(string)));
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", "Header", "A1", 11, 6)]
        [DataRow("BIFF8_data.xls", "Header", "A1", 11, 6)]

        public void ReadRange_WithoutHeaderAll_ReturnRange(string fileName, string sheetName, string range, int expectedRowCount, int expectedColumnCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange(sheetName, range, false, true);
            workbook.Dispose();
            Assert.AreEqual(expectedRowCount, dataTable.Rows.Count);
            Assert.AreEqual(expectedColumnCount, dataTable.Columns.Count);

            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.DataType);
            Assert.IsTrue(dataTable.Columns.Contains("Col1"));
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", "HeadersOnly", true, true)]
        [DataRow("BIFF8_data.xls", "HeadersOnly", true, true)]
        [DataRow("OXML_data.xlsx", "HeadersOnly", false, false)]
        [DataRow("BIFF8_data.xls", "HeadersOnly", false, false)]
        public void ReadRange_HeadersOnly_ReturnEmptyTableWithColumns(string fileName, string sheetName, bool hasHeaders, bool useColumnDataType)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dt = workbook.ReadRange(sheetName, "A1", hasHeaders, useColumnDataType);
            workbook.Dispose();
            Assert.AreEqual(3, dt.Columns.Count);
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1:C4", 3, 3, true, true)]
        [DataRow("OXML_data.xlsx", "E3:I6", 3, 5, true, true)]
        [DataRow("OXML_data.xlsx", "A1:C4", 4, 3, false, false)]
        [DataRow("OXML_data.xlsx", "E3:I6", 4, 5, false, false)]
        [DataRow("OXML_data.xlsx", "B9:F12", 3, 5, true, false)]
        [DataRow("OXML_data.xlsx", "B9:F12", 4, 5, false, false)]
        public void ReadRange_Tables_ReturnRange(string fileName, string range, int expectedRows, int expectedCols, bool hasHeaders, bool useColumnDataType)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dt = workbook.ReadRange("Tables", range, hasHeaders, useColumnDataType);
            workbook.Dispose();
            Assert.AreEqual(expectedRows, dt.Rows.Count);
            Assert.AreEqual(expectedCols, dt.Columns.Count);
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx")]
        public void ReadRange_EmptyColData_ReturnRange(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dt = workbook.ReadRange("EmptyColData", "A1", false, false);
            workbook.Dispose();
            Assert.AreEqual(5, dt.Columns.Count);
        }


        [TestMethod]
        public void ReadRange_EmtpySheet_ReturnsEmptyDataTable()
        {
            var path = IOSamples.GetSamplePath("OXML_data.xlsx");
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dt = workbook.ReadRange("EmptySheet", "A1", true, false);
            workbook.Dispose();
            Assert.IsNotNull(dt);
            Assert.AreEqual(0, dt.Rows.Count);
            Assert.AreEqual(0, dt.Columns.Count);
        }
    }
}
