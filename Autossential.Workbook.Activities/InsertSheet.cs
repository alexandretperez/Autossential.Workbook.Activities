using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class InsertSheet : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        public InArgument<int?> Position { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var position = Position.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException();

            context.GetWorkbookProcessor().InsertSheet(sheetName, position);
        }
    }
}
