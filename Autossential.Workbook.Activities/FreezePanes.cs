using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class FreezePanes : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        
        [RequiredArgument]
        public InArgument<int> ColumnsToFreeze { get; set; } = 0;

        [RequiredArgument]
        public InArgument<int> RowsToFreeze { get; set; } = 1;
        
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.FreezePanes_SheetName_DisplayName));

            var colsToFreeze = ColumnsToFreeze.Get(context);
            var rowsToFreeze = RowsToFreeze.Get(context);

            if (colsToFreeze < 0)
                colsToFreeze = 0;

            if (rowsToFreeze < 0)
                rowsToFreeze = 0;

            context.GetWorkbookProcessor().FreezePanes(sheetName, colsToFreeze, rowsToFreeze);
        }
    }
}
