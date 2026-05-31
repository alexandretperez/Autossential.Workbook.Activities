using Autossential.Workbook.Activities.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Extensions
{
    public static class ActivityContextExtensions
    {
        public const string WorkbookInstancePropertyName = "_wbInst";

        extension(ActivityContext context)
        {
            public IWorkbookProcessor GetWorkbookProcessor()
            {
                return context.DataContext.GetProperties()[WorkbookInstancePropertyName]?.GetValue(context.DataContext) as IWorkbookProcessor;
            }
        }
    }
}