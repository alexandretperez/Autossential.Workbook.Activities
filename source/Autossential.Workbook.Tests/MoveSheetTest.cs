using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Adapters;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class MoveSheetTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            IOSamples.ExportSample("empty.xlsx", "merge.xlsx");
            IOSamples.ExportSample("empty.xls", "merge.xls");
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("merge.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("merge.xls"));
        }

        private void GenerateData(IWorkbookAdapter adapter)
        {
            adapter.WriteCell("Sheet1", "B2", 1);
            adapter.WriteCell("Sheet4", "B2", 4);
            adapter.WriteCell("Sheet5", "B2", 5);
        }

        [TestMethod]
        public void OpenXML()
        {
            var path = IOSamples.GetTestPath("merge.xlsx");
            var adapter = WorkbookAdapterFactory.Create(path);
            GenerateData(adapter);
           
            adapter.MoveSheet("Sheet2", 1); // 1,2,4,5
            adapter.MoveSheet("Sheet1", 2, true, "Sheet3"); // 1,2,3,4,5
            adapter.Save();
            adapter.Dispose();
        }


        [TestMethod]
        public void OLE2()
        {
            var path = IOSamples.GetTestPath("merge.xls");
            var adapter = WorkbookAdapterFactory.Create(path);
            GenerateData(adapter);

            adapter.MoveSheet("Sheet2", -3); // 1,2,4,5
            adapter.MoveSheet("Sheet6", -1); // 1,2,4,5,6
            adapter.MoveSheet("Sheet3", -4); // 1,2,3,4,5
            adapter.Save();
            adapter.Dispose();
        }
    }
}