using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class RenameSheet : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<string> NewSheetName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var newSheetName = NewSheetName.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException();

            if (string.IsNullOrEmpty(newSheetName))
                throw new InvalidOperationException();

            context.GetWorkbookProcessor().RenameSheet(sheetName, newSheetName);
        }
    }
}
