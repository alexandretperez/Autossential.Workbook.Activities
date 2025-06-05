using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{
    [TestClass]
    public class ActivateSheet_Tests
    {
        [TestMethod]
        [DataRow("BIFF8_sheets.xls", "Sheet1")]
        [DataRow("BIFF8_sheets.xls", "Sheet2")]
        [DataRow("BIFF8_sheets.xls", "Sheet3")]
        [DataRow("BIFF8_sheets.xls", "Sheet4")]

        [DataRow("OXML_sheets.xlsx", "Sheet1")]
        [DataRow("OXML_sheets.xlsx", "Sheet2")]
        [DataRow("OXML_sheets.xlsx", "Sheet3")]
        [DataRow("OXML_sheets.xlsx", "Sheet4")]
        public void ActivateSheet_SheetNameExists_SheetActivated(string fileName, string sheetName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            workbook.ActivateSheet(sheetName);
            Assert.AreEqual(sheetName, workbook.GetActiveSheet().name);

            workbook.Save();
            workbook.Dispose();

            workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.AreEqual(sheetName, workbook.GetActiveSheet().name);
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        [ExpectedException(typeof(ArgumentException))]
        public void ActivateSheet_SheetNameDoesNotExist_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            workbook.ActivateSheet("DoesNotExist");
            workbook.Dispose();
        }

        //[TestMethod]
        //public void Test()
        //{
        //    var path = @"D:\Users\alexa\Downloads\Financial Sample.xlsx";
        //    var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
        //    workbook.ActivateSheet("Sheet3");
        //    workbook.Dispose();
        //}
    }
}
