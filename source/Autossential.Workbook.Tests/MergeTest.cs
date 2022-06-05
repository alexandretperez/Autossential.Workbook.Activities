using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Adapters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class MergeTest
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

        [TestMethod]
        public void OpenXML()
        {
            var path = IOSamples.GetTestPath("merge.xlsx");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.MergeRange("Sheet1", "B2:F2");
            adapter.MergeRange("Sheet1", "B4:B10");
            adapter.MergeRange("Sheet1", "D4:F10");
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        public void OLE2()
        {
            var path = IOSamples.GetTestPath("merge.xls");
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.MergeRange("Sheet1", "B2:F2");
            adapter.MergeRange("Sheet1", "B4:B10");
            adapter.MergeRange("Sheet1", "D4:F10");
            adapter.Save();
            adapter.Dispose();
        }
    }
}