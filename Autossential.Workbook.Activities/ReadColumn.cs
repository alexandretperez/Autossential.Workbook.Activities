using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadColumn : WorkbookCodeActivity<object[]>
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<string> StartingCell { get; set; }
        public InArgument<int> Limit { get; set; }
        protected override object[] Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var startingCell = StartingCell.Get(context);
            var limit = Limit.Get(context);

            return context.GetWorkbookProcessor().ReadColumn(sheetName, startingCell, limit);
        }
    }
}
