using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.PowerFx.Core.Public.Values;
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
        [DataRow("OXML_data.xlsx", "Header", "B4", 7, 5)]

        [DataRow("BIFF8_data.xls", "Header", "A1", 10, 6)]
        [DataRow("BIFF8_data.xls", "Header", "B1", 10, 5)]
        [DataRow("BIFF8_data.xls", "Header", "B4", 7, 5)]

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
        [DataRow("OXML_data.xlsx", "Header", "A1", 11, 6)]
        [DataRow("OXML_data.xlsx", "Header", "B1", 11, 5)]
        [DataRow("OXML_data.xlsx", "Header", "B4", 8, 5)]

        [DataRow("OXML_data.xlsx", "NoHeader", "A1", 10, 6)]
        [DataRow("OXML_data.xlsx", "NoHeader", "B1", 10, 5)]
        [DataRow("OXML_data.xlsx", "NoHeader", "B4", 7, 5)]

        [DataRow("BIFF8_data.xls", "Header", "A1", 11, 6)]
        [DataRow("BIFF8_data.xls", "Header", "B1", 11, 5)]

        // BUG: https://github.com/MarkPflug/Sylvan/issues/267
        [DataRow("BIFF8_data.xls", "Header", "B4", 8, 5)]
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
                    typeof(int),
                    typeof(decimal),
                    typeof(decimal),
                    typeof(DateTime),
                    typeof(string),
                    typeof(bool)
                });
            }
            else
            {
                Assert.IsTrue(columnTypes.All(p => p == typeof(string)));
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

            Assert.IsTrue(columnTypes.All(type => type == typeof(bool)));
        }
    }
}
