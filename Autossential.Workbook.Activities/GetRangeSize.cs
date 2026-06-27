using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetRangeSize : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public OutArgument<int> RowCount { get; set; }
        public OutArgument<int> ColumnCount { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.GetRangeSize_SheetName_DisplayName));

            if (string.IsNullOrEmpty(range))
                range = "A1";

            var workbookProcessor = context.GetWorkbookProcessor();
            RowCount?.Set(context, workbookProcessor.GetRowCount(sheetName, range));
            ColumnCount?.Set(context, workbookProcessor.GetColumnCount(sheetName, range));
        }
    }
}
