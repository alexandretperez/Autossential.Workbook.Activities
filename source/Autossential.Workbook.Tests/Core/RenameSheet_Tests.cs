using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{

    [TestClass]
    public class RenameSheet_Tests
    {
        [TestMethod]
        [DataRow("OXML_data.xlsx")]
        [DataRow("BIFF8_data.xls")]
        public void RenameSheet_SheetIndexValid_SheetNameChanged(string fileName)
        {
            // Arrange
            var path = IOSamples.GetSamplePath(fileName);
            const string expectedSheetName = "My Data";
            const int sheetIndex = 0;

            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);

            // Act
            var originalSheetName = workbook.GetSheetNames()[0];
            workbook.RenameSheet(sheetIndex, expectedSheetName);
            workbook.Save();
            workbook.Dispose();

            workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var sheetName = workbook.GetSheetNames()[0];
            Assert.AreEqual(expectedSheetName, sheetName);

            workbook.RenameSheet(0, originalSheetName);
            workbook.Save();
            workbook.Dispose();
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void RenameSheet_SheetIndexInvalid_ThrowsException()
        {
            // Arrange
            // Arrange
            var path = IOSamples.GetSamplePath("OXML_data.xlsx");
            const string expectedSheetName = "NewSheetName";
            const int invalidSheetIndex = 999; // Assuming this index doesn't exist

            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);

            // Act
            workbook.RenameSheet(invalidSheetIndex, expectedSheetName);
            workbook.Save();
            workbook.Dispose();
        }
    }
}
