using Autossential.Workbook.Activities.Core;

namespace Autossential.Workbook.Activities.Tests.Unit
{

    [ClassDataSource<SharedSource>(Shared = SharedType.PerClass)]
    public class WorkbookProcessorTest(SharedSource source)
    {
        private readonly SharedSource source = source;

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Dispose_WhenNotWrites_DoesNotCreateFile(string extension)
        {
            var (processor, filePath) = source.NewFile(extension);
            processor.Dispose();
            await Assert.That(File.Exists(filePath)).IsFalse();
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Dispose_WhenWrites_CreateFile(string extension)
        {
            var (processor, filePath) = source.NewFile(extension);
            processor.WriteCell("Sheet1", "A1", "Hello");
            processor.Dispose();
            await Assert.That(File.Exists(filePath)).IsTrue();
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Save_WhenMultipleCalls_IgnoresSubsequent(string extension)
        {
            var (processor, filePath) = source.NewFile(extension);

            processor.Save();
            var first = File.GetLastWriteTime(filePath);
            await Task.Delay(50);

            processor.Save();
            var second = File.GetLastWriteTime(filePath);
            await Task.Delay(50);

            processor.Dispose();
            var third = File.GetLastWriteTime(filePath);

            await Assert.That(first).IsEqualTo(second).And.IsEqualTo(third);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task WriteRange_MissingSheet_CreateNewSheetAndWrite(string extension)
        {
            var (processor, _) = source.NewFile(extension);
            var data = TableUtils.Generate(4, 5, "a");
            processor.WriteRange("New Sheet", data, "A1", true);

            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["Sheet1", "New Sheet"]);
            var readData = processor.ReadRange("New Sheet", "A1", true, 1, 1);

            await Assert.That(TableUtils.AreTablesEqual(data, readData)).IsTrue();
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task ReadRange_ExistingFile_ReadSheet(string extension)
        {
            var (processor, filePath) = source.NewFile(extension);
            var data = TableUtils.Generate(5, 10, "b");
            processor.WriteRange("Sheet1", data, "A1", true);
            processor.Dispose();

            processor = WorkbookProcessorFactory.OpenOrCreate(filePath);

            var readData = processor.ReadRange("Sheet1", "A1", true, 1, 1);
            await Assert.That(TableUtils.AreTablesEqual(data, readData)).IsTrue();
        }

        [Test]
        [NotInParallel]
        [Arguments(".xlsx", "Sheet1", "A1", 7, 7)]
        [Arguments(".xls", "Sheet2", "A1", 10, 10)]
        [Arguments(".xlsx", "Sheet3", "A1", 4, 11)]

        [Arguments(".xlsx", "Sheet1", "B3:F6", 5, 3)]
        [Arguments(".xls", "Sheet2", "E2", 4, 9)]
        [Arguments(".xlsx", "Sheet2", "C7", 6, 4)]
        public async Task ReadRange_InMemory_ReadSheet(string extension, string sheet, string range, int expectedCols, int expectedRows)
        {
            var (processor, _) = source.SharedFile(extension);
            var readData = processor.ReadRange(sheet, range, true, 1, 1);
            await Assert.That(expectedRows).IsEqualTo(readData.Rows.Count);
            await Assert.That(expectedCols).IsEqualTo(readData.Columns.Count);
        }

        [Test]
        [Arguments(3)]
        public async Task ReadRange_HeaderAndRecordRows_ReadSheet(int expectedRows)
        {
            var (processor, _) = source.SharedFile(".xlsx");
            var readData = processor.ReadRange("Sheet3", "A2", true, 2, 3);
            _ = await Assert.That(expectedRows).IsEqualTo(readData.Rows.Count);
        }
    }
}
