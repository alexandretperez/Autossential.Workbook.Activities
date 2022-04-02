using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class ReadRangeTest
    {
        [TestMethod]
        //[DataRow(@"D:\Users\alexa\Downloads\Book1.xlsx", "C1")]
        [DataRow(@"D:\Users\alexa\Downloads\Book1.xlsx", "")]
        public void Read(string file, string address)
        {
            var path = file; //IOSamples.GetSamplePath("book.xlsx");
            var adapter = WorkbookAdapterFactory.Create(path);
            var dt = adapter.ReadRangeAsync("sheet1", address, true);
            Assert.IsNotNull(dt);
        }
    }
}