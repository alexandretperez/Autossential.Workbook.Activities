using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class DeleteRows : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<string> RowReferences { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.DeleteRows_SheetName_DisplayName));

            var rowReferences = RowReferences.Get(context);
            if (string.IsNullOrEmpty(rowReferences))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.DeleteRows_RowReferences_DisplayName));

            context.GetWorkbookProcessor().DeleteRows(sheetName, rowReferences);
        }
    }
}
