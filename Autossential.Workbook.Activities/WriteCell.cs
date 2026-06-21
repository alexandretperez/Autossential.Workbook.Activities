using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class WriteCell : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        [RequiredArgument]
        public InArgument<string> CellAddress { get; set; }

        [RequiredArgument]
        public InArgument<object> Value { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var cell = CellAddress.Get(context);
            var value = Value.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.WriteCell_SheetName_DisplayName));

            if (string.IsNullOrEmpty(cell))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.WriteCell_CellAddress_DisplayName));

            context.GetWorkbookProcessor().WriteCell(sheetName, cell, value);
        }
    }
}
