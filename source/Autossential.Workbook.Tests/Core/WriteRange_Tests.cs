using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{
    [TestClass]
    public class WriteRange_Tests
    {
        [TestMethod]
        public void WriteRange_Sheet1()
        {
            var path = @"D:\Users\alexa\Downloads\NewFile.xlsx";
            var dt = TableGenerator.Generate(10, 10);

            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            workbook.WriteRange(dt, "Sheet1", "A1", true);
            workbook.Save();
        }

        [TestMethod]
        public void WriteRange_Sheet2()
        {
            var path = @"D:\Users\alexa\Downloads\NewFile.xlsx";
            var dt = TableGenerator.Generate(10, 10);

            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            workbook.WriteRange(dt, "Sheet2", "A1", true);
            workbook.Save();
        }

        [TestMethod]
        public void WriteRange_Sheet1_Update()
        {
            var path = @"D:\Users\alexa\Downloads\NewFile.xlsx";
            var dt = TableGenerator.Generate(10, 10);

            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            workbook.WriteRange(dt, "Sheet1", "D4", true);
            workbook.Save();
        }

    }
}
