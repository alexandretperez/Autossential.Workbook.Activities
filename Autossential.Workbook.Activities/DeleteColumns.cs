using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class DeleteColumns : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<string> ColumnReferences { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.DeleteColumns_SheetName_DisplayName));

            var columnReferences = ColumnReferences.Get(context);
            if (string.IsNullOrEmpty(columnReferences))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.DeleteColumns_ColumnReferences_DisplayName));

            context.GetWorkbookProcessor().DeleteColumns(sheetName, columnReferences);
        }
    }
}
