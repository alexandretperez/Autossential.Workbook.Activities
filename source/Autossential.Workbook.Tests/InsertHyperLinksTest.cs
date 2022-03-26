using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class InsertHyperLinksTest
    {
        private static string _openXmlFile;
        private static string _ole2file;

        [ClassInitialize]
        public static void ClassInitialize(TestContext _)
        {
            IOSamples.ClearFolder();
            _openXmlFile = IOSamples.ExportSample("book.xlsx");
            _ole2file = IOSamples.ExportSample("book.xls");
        }

        [ClassCleanup]
        public static void CleanUp()
        {
            IOSamples.ClearFolder();
        }

        [TestMethod]
        [DataRow("A1", "https://www.google.com/", "")] // with '/' on the end
        [DataRow("B2", "mailto:email@email.com", "")]
        [DataRow("C3", "../sample.txt", "")]
        [DataRow("D4", "Sheet2!E10", "")]
        [DataRow("E5", "ftp://some.ftp.com", "")] // without '/' on the end
        [DataRow("A7", "https://www.google.com/", "Google")] // with '/' on the end
        [DataRow("A8", "mailto:email@email.com", "Email")]
        [DataRow("A9", "../sample.txt", "Sample")]
        [DataRow("A10", "Sheet2!E10", "Reference")]
        [DataRow("A11", "ftp://some.ftp.com", "FTP")] // without '/' on the end
        public void OpenXML(string cell, string link, string label)
        {
            var args = CreateInsertHyperlinkArgs(_openXmlFile, cell, link, label);
            var workflow = new InsertHyperlink { UseScope = false };
            var result = WorkflowTester.Run(workflow, args);
            var value = result.Get(p => p.Result);
            Assert.AreEqual(value, true);
        }

        [TestMethod]
        [DataRow("A1", "https://www.google.com/", "")] // with '/' on the end
        [DataRow("B2", "mailto:email@email.com", "")]
        [DataRow("C3", "../sample.txt", "")]
        [DataRow("D4", "Sheet2!E10", "")]
        [DataRow("E5", "ftp://some.ftp.com", "")] // without '/' on the end
        [DataRow("A7", "https://www.google.com/", "Google")] // with '/' on the end
        [DataRow("A8", "mailto:email@email.com", "Email")]
        [DataRow("A9", "../sample.txt", "Sample")]
        [DataRow("A10", "Sheet2!E10", "Reference")]
        [DataRow("A11", "ftp://some.ftp.com", "FTP")] // without '/' on the end
        public void OLE2(string cell, string link, string label)
        {
            var args = CreateInsertHyperlinkArgs(_ole2file, cell, link, label);
            var workflow = new InsertHyperlink { UseScope = false };
            var result = WorkflowTester.Run(workflow, args);
            Assert.AreEqual(result.Get(p => p.Result), true);
        }

        private static Dictionary<string, object> CreateInsertHyperlinkArgs(string file, string cell, string link, string label)
        {
            return new Dictionary<string, object>
            {
                { nameof(InsertHyperlink.WorkbookPath), file },
                { nameof(InsertHyperlink.SheetName), "Sheet1" },
                { nameof(InsertHyperlink.Cell), cell},
                { nameof(InsertHyperlink.Label), label},
                { nameof(InsertHyperlink.Link), link},
                { nameof(InsertHyperlink.Tooltip), "My link address is " + link } // OpenXML only
            };
        }
    }
}
