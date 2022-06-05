using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class RenameSheetTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            GenerateData(IOSamples.ExportSample("empty.xlsx", "rename.xlsx"));
            GenerateData(IOSamples.ExportSample("empty.xls", "rename.xls"));
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("rename.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("rename.xls"));
        }

        private static void GenerateData(string path)
        {
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.WriteCell("Sheet2", "A1", "");
            adapter.Save();
            adapter.Dispose();
        }


        [TestMethod]
        public void OpenXML()
        {
            var path = IOSamples.GetTestPath("rename.xlsx");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.RenameSheet(0, "Autossential");
            adapter.RenameSheet(1, "Activities");

            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        public void OLE2()
        {
            var path = IOSamples.GetTestPath("rename.xls");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.RenameSheet(0, "Autossential");
            adapter.RenameSheet(1, "Activities");

            adapter.Save();
            adapter.Dispose();
        }
    }
}