using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Autossential.Workbook.Activities.Tests
{
    public static class WorkbookGenerator
    {
        public static string CreateWorkbookFile(bool openXmlFormat, params Action<ISheet>[] config)
        {
            IWorkbook wb = openXmlFormat ? new XSSFWorkbook() : new HSSFWorkbook();
            var extension = openXmlFormat ? ".xlsx" : ".xls";
            var path = Path.Combine(Path.GetTempPath(), $"workbook_{Guid.NewGuid():N}{extension}");
            var index = 1;
            foreach (var cfg in config)
            {
                var sheet = wb.CreateSheet($"Sheet{index++}");
                cfg(sheet);
            }
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fs);
            return path;
        }
    }
}
