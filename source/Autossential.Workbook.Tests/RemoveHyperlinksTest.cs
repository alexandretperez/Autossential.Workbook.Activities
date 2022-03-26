using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Adapters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class RemoveHyperlinksTest
    {
        private static string _openXmlFile;
        private static string _ole2file;

        [TestInitialize]
        public void Initialize()
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
        [DataRow(1, "A1")]
        [DataRow(2, "A1:B2")]
        [DataRow(3, "A1:C3")]
        [DataRow(4, "A1:D4")]
        [DataRow(5, "A1:E5")]
        [DataRow(5, "")]

        public async Task OLE2(int expectedRemovedLinks, string cellRange)
        {
            InsertHyperlinks(_ole2file);
            var args = CreateRemoveHyperlinksArgs(_ole2file, cellRange);
            var workflow = new RemoveHyperlinks { UseScope = false };
            var result = WorkflowTester.Run(workflow, args);
            Assert.AreEqual(expectedRemovedLinks, result.Get(p => p.Result));
            using (var adapter = new OLE2WorkbookAdapter(_ole2file))
            {
                var links = await adapter.GetHyperlinksAsync("Sheet1", "");
                Assert.AreEqual(5 - expectedRemovedLinks, links.Length, "remaining links differ");
            }
        }

        [TestMethod]
        [DataRow(1, "A1")]
        [DataRow(2, "A1:B2")]
        [DataRow(3, "A1:C3")]
        [DataRow(4, "A1:D4")]
        [DataRow(5, "A1:E5")]
        [DataRow(5, "")]

        public void OpenXML(int expectedResult, string cellRange)
        {
            InsertHyperlinks(_openXmlFile);
            var args = CreateRemoveHyperlinksArgs(_openXmlFile, cellRange);
            var workflow = new RemoveHyperlinks { UseScope = false };
            var result = WorkflowTester.Run(workflow, args);

            Assert.AreEqual(expectedResult, result.Get(p => p.Result));
        }

        private void InsertHyperlinks(string file)
        {
            var dic = new Dictionary<string, string>
                {
                    { "A1", "https://www.url.com/" }, // with '/' on the end
                    { "B2", "mailto:email@email.com" },
                    { "C3", "../sample.txt" },
                    { "D4", "Sheet2!E10" },
                    { "E5", "ftp://some.ftp.com" } // without '/' on the end
                };

            foreach (var item in dic)
            {
                var workflow = new InsertHyperlink
                {
                    UseScope = false,
                    WorkbookPath = file,
                    Cell = item.Key,
                    Link = item.Value,
                    SheetName = "Sheet1"
                };

                WorkflowTester.Invoke(workflow);
            }
        }

        private static Dictionary<string, object> CreateRemoveHyperlinksArgs(string file, string cellRange)
        {
            return new Dictionary<string, object>
            {
                { nameof(RemoveHyperlinks.WorkbookPath), file },
                { nameof(RemoveHyperlinks.SheetName), "Sheet1" },
                { nameof(RemoveHyperlinks.Range), cellRange}
            };
        }
    }
}
