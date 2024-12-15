using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class RenameSheet : WorkbookCodeActivity
    {
        public InArgument<int> SheetIndex { get; set; }

        public InArgument<string> FromSheetName { get; set; }

        [RequiredArgument]
        public InArgument<string> ToSheetName { get; set; }

        public bool FindByIndex { get; set; } = false;

        protected override void Execute(CodeActivityContext context)
        {
            var sheetIndex = SheetIndex.Get(context);
            var fromSheetName = FromSheetName.Get(context);
            var toSheetName = ToSheetName.Get(context);

            var workbook = context.GetWorkbook();
            if (FindByIndex)
                workbook.RenameSheet(sheetIndex, toSheetName);
            else
                workbook.RenameSheet(fromSheetName, toSheetName);
        }
    }
}
