using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Enums;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class DrawBorderTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            IOSamples.ExportSample("empty.xlsx", "borders.xlsx");
            IOSamples.ExportSample("empty.xls", "borders.xls");
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("borders.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("borders.xls"));
        }

        [TestMethod]
        [DataRow(Border.All, "borders.xlsx", "B2:D7")]
        [DataRow(Border.Top, "borders.xlsx", "F2:H7")]
        [DataRow(Border.Bottom, "borders.xlsx", "J2:L7")]
        [DataRow(Border.Left, "borders.xlsx", "N2:P7")]
        [DataRow(Border.Right, "borders.xlsx", "R2:T7")]
        [DataRow(Border.Outside, "borders.xlsx", "B9:D14")]
        [DataRow(Border.Inside, "borders.xlsx", "F9:H14")]

        public void OpenXml(Border border, string fileName, string range)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.DrawBorder("Sheet1", range, border, BorderStyle.Thick, Color.Blue);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow(Border.All, "borders.xls", "B2:D7")]
        [DataRow(Border.Top, "borders.xls", "F2:H7")]
        [DataRow(Border.Bottom, "borders.xls", "J2:L7")]
        [DataRow(Border.Left, "borders.xls", "N2:P7")]
        [DataRow(Border.Right, "borders.xls", "R2:T7")]
        [DataRow(Border.Outside, "borders.xls", "B9:D14")]
        [DataRow(Border.Inside, "borders.xls", "F9:H14")]
        public void OLE2(Border border, string fileName, string range)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.DrawBorder("Sheet1", range, border, BorderStyle.Thick, Color.Blue);
            adapter.Save();
            adapter.Dispose();
        }
    }
}
