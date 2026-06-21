using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class SaveWorkbook : WorkbookCodeActivity
    {
        protected override void Execute(CodeActivityContext context)
        {
            context.GetWorkbookProcessor().Save();
        }
    }
}
