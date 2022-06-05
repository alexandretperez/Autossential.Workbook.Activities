using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class DeleteSheetTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            GenerateData(IOSamples.ExportSample("empty.xlsx", "delete.xlsx"));
            GenerateData(IOSamples.ExportSample("empty.xls", "delete.xls"));
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("delete.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("delete.xls"));
        }

        private static void GenerateData(string path)
        {
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.WriteCell("Sheet2", "A1", "");
            adapter.WriteCell("Sheet3", "A1", "");
            adapter.WriteCell("Sheet4", "A1", "");
            adapter.Save();
            adapter.Dispose();
        }


        [TestMethod]
        public void OpenXML()
        {
            var path = IOSamples.GetTestPath("delete.xlsx");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.DeleteSheet("Sheet2");
            adapter.DeleteSheet("Sheet3");
            adapter.DeleteSheet("NonExistingSheet");

            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        public void OLE2()
        {
            var path = IOSamples.GetTestPath("delete.xls");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.DeleteSheet("Sheet2");
            adapter.DeleteSheet("Sheet3");
            adapter.DeleteSheet("NonExistingSheet");

            adapter.Save();
            adapter.Dispose();
        }
    }
}