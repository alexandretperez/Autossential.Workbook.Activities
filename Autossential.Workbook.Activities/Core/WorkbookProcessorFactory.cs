using Autossential.Workbook.Activities.Core.Processors;
using System.Text;

namespace Autossential.Workbook.Activities.Core
{
    public static class WorkbookProcessorFactory
    {
        public static IWorkbookProcessor OpenOrCreate(string filePath, string password = null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".xlsx" or ".xlsb" => new OpenXMLWorkbookProcessor(filePath, password),
                ".xls" => new BinaryWorkbookProcessor(filePath, password),
                _ => throw new NotSupportedException($"File type '{extension}' is not supported. Only .xlsx and .xls are supported.")
            };
        }
    }
}
