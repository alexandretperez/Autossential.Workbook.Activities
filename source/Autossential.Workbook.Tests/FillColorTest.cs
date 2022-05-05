using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
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
            IOSamples.ExportSample("book.xlsx");
            IOSamples.ExportSample("book.xls");
        }

        [TestMethod]
        [DataRow("book.xlsx")]
        [DataRow("book.xls")]
        public void HorizontalFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Horizontal", "A1:E5", new[] { Color.Teal, Color.Orange }, Core.Enums.FillOrientation.Horizontal);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("book.xlsx")]
        [DataRow("book.xls")]
        public void VerticalFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Vertical", "A1:E5", new[] { Color.Teal, Color.Orange }, Core.Enums.FillOrientation.Vertical);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("book.xlsx")]
        [DataRow("book.xls")]
        public void ChessFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Chess", "A1:E5", new[] { Color.Teal, Color.Orange }, Core.Enums.FillOrientation.Chess);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("book.xlsx")]
        [DataRow("book.xls")]
        public void DiagonalLeft(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("DiagonalLeft", "A1:E5", new[] { Color.Teal, Color.Orange }, Core.Enums.FillOrientation.DiagonalLeft);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("book.xlsx")]
        [DataRow("book.xls")]
        public void DiagonalRight(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("DiagonalRight", "A1:E5", new[] { Color.Teal, Color.Orange }, Core.Enums.FillOrientation.DiagonalRight);
            adapter.Save();
            adapter.Dispose();
        }
    }
}