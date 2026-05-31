using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
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
            return context.GetWorkbookProcessor().ReadCell(sheetName, cellAddress);
        }
    }
}
