using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetNames : WorkbookCodeActivity<string[]>
    {
        protected override string[] Execute(CodeActivityContext context)
        {
            var wb = context.GetWorkbook();
            return wb.GetSheetNames();
        }
    }
}