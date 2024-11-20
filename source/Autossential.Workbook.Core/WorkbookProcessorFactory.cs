using Autossential.Workbook.Core.Processors;
using System;
using System.IO;
using System.Text;

namespace Autossential.Workbook.Core
{
    public static class WorkbookProcessorFactory
    {
        private static void RegisterProvider()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static IWorkbookProcessor OpenOrCreate(string path, bool createIfNotExists = false)
        {
            RegisterProvider();

            var extension = Path.GetExtension(path).ToLowerInvariant();
            switch (extension)
            {
                case ".xlsm": // macro enabled workbook
                case ".xltm": // macro enabled template
                case ".xlsx": // workbook
                case ".xltx": // template
                    return new OpenXMLWorkbookProcessor(path, createIfNotExists);
                case ".xls": // workbook
                    return new BIFF8WorkbookProcessor(path, createIfNotExists);
                default:
                    throw new InvalidOperationException("The file stream must be a BIFF8 or OOXML stream");
            }
        }
    }
}
