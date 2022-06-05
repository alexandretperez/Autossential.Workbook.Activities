using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Enums;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class FillColorTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            IOSamples.ExportSample("empty.xlsx", "colors.xlsx");
            IOSamples.ExportSample("empty.xls", "colors.xls");
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("colors.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("colors.xls"));
        }


        [TestMethod]
        [DataRow(FillOrientation.Horizontal, "colors.xlsx", "B2:D7")]
        [DataRow(FillOrientation.Vertical, "colors.xlsx", "F2:H7")]
        [DataRow(FillOrientation.DiagonalBackward, "colors.xlsx", "J2:L7")]
        [DataRow(FillOrientation.DiagonalForward, "colors.xlsx", "N2:P7")]
        [DataRow(FillOrientation.Chess, "colors.xlsx", "R2:T7")]

        public void OpenXml(FillOrientation orientation, string fileName, string range)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Sheet1", range, new[] { Color.Teal, Color.Orange }, orientation);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow(FillOrientation.Horizontal, "colors.xls", "B2:D7")]
        [DataRow(FillOrientation.Vertical, "colors.xls", "F2:H7")]
        [DataRow(FillOrientation.DiagonalBackward, "colors.xls", "J2:L7")]
        [DataRow(FillOrientation.DiagonalForward, "colors.xls", "N2:P7")]
        [DataRow(FillOrientation.Chess, "colors.xls", "R2:T7")]

        public void OLE2(FillOrientation orientation, string fileName, string range)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Sheet1", range, new[] { Color.Teal, Color.Orange }, orientation);
            adapter.Save();
            adapter.Dispose();
        }

    }
}