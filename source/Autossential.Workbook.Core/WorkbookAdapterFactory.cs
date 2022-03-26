using Autossential.Workbook.Core.Adapters;
using System;
using System.IO;

namespace Autossential.Workbook.Core
{
    public static class WorkbookAdapterFactory
    {
        public static IWorkbookAdapter Create(string path)
        {
            IWorkbookAdapter adapter = null;

            var extension = Path.GetExtension(path);
            switch (extension.ToLowerInvariant())
            {
                case ".xlsm": // macro enabled workbook
                case ".xltm": // macro enabled template
                case ".xlsx": // workbook
                case ".xltx": // template
                    adapter = new OpenXmlWorkbookAdapter(path);
                    break;
                default:
                    adapter = new OLE2WorkbookAdapter(path);
                    break;
            }

            if (adapter == null)
                throw new InvalidOperationException("The file stream need be a OLE2 or OOXML stream");

            return adapter;
        }
    }
}
