using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class DeleteSheet : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException();

            context.GetWorkbookProcessor().DeleteSheet(sheetName);
        }
    }
}
