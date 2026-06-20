using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadCell : WorkbookCodeActivity<object>
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        [RequiredArgument]
        public InArgument<string> CellAddress { get; set; }
        protected override object Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var cellAddress = CellAddress.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.ReadCell_SheetName_DisplayName));

            if (string.IsNullOrEmpty(cellAddress))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.ReadCell_CellAddress_DisplayName));

            return context.GetWorkbookProcessor().ReadCell(sheetName, cellAddress);
        }
    }
}
