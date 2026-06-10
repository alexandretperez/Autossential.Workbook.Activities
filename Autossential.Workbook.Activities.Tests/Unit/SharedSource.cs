using Autossential.Workbook.Activities.Core;
using TUnit.Core.Interfaces;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class SharedSource : IAsyncInitializer, IAsyncDisposable
    {
        public ValueTask DisposeAsync()
        {
            foreach (var item in _cache)
            {
                var (processor, filePath) = item.Value;
                processor.Dispose();
                File.Delete(filePath);
            }

            foreach (var filePath in _newFiles)
                File.Delete(filePath);

            return ValueTask.CompletedTask;
        }

        private readonly Dictionary<string, (IWorkbookProcessor, string)> _cache = [];
        private readonly List<string> _newFiles = [];
        public (IWorkbookProcessor Processor, string FilePath) NewFile(string extension)
        {
            var filePath = Path.ChangeExtension(Path.GetTempFileName(), extension);
            _newFiles.Add(filePath);
            return (WorkbookProcessorFactory.OpenOrCreate(filePath), filePath);
        }

        public (IWorkbookProcessor Processor, string FilePath) SharedFile(string extension) =>
            _cache[extension];

        public Task InitializeAsync()
        {
            var sheet1 = TableUtils.Build(7, 7, (col, row) => col == row ? $"c{col}r{row}" : null);
            var sheet2 = TableUtils.Build(10, 10, (col, row) =>
            {
                if ((col == 1 || col % 4 == 0) && row < 5)
                    return $"c{col}r{row}";

                if ((col % 2 == 0 && col < 10) && row > 5)
                    return $"c{col}r{row}";

                return null;
            });
            var sheet3 = TableUtils.Build(4, 11, (col, row) =>
            {
                if (row < 3)
                    return $"h{col}.{row}";

                return $"c{col}r{row}";
            });

            foreach (var extension in new string[] { ".xlsx", ".xls" })
            {
                var filePath = Path.ChangeExtension(Path.GetTempFileName(), extension);
                var processor = WorkbookProcessorFactory.OpenOrCreate(filePath);
                processor.WriteRange("Sheet1", sheet1, "A1", true);
                processor.WriteRange("Sheet2", sheet2, "A1", true);
                processor.WriteRange("Sheet3", sheet3, "A1", true);
                _cache.Add(extension, (processor, filePath));
            }

            return Task.CompletedTask;
        }
    }
}
