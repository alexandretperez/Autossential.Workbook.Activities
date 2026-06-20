using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class FindValue : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Range { get; set; }
        [RequiredArgument]
        public InArgument<object> Value { get; set; }
        public OutArgument<string> CellAddress { get; set; }
        public OutArgument<int> RowNumber { get; set; }
        public OutArgument<int> ColumnNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context) ?? "A1";
            var value = Value.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.FindValue_SheetName_DisplayName));

            var (address, col, row) = context.GetWorkbookProcessor().FindValue(sheetName, range, value);

            CellAddress.Set(context, address);
            ColumnNumber.Set(context, col);
            RowNumber.Set(context, row);
        }
    }
}
