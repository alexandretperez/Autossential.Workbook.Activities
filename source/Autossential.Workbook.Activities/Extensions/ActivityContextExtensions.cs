using Autossential.Workbook.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Extensions
{
    public static class ActivityContextExtensions
    {
        public const string WorkbookInstancePropertyName = "WorkbookInstance";
        public static IWorkbookProcessor GetWorkbook(this ActivityContext context)
        {
            return context.DataContext.GetProperties()[WorkbookInstancePropertyName]?.GetValue(context.DataContext) as IWorkbookProcessor;
        }
    }
}