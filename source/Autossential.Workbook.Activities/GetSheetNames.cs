using Autossential.Shared.Activities.Constraints;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetNames : WorkbookCodeActivity<string[]>
    {
        protected override string[] Execute(CodeActivityContext context)
        {
            var wb = context.GetWorkbook();
            try
            {
                return wb.GetSheetNames();
            }
            finally
            {
                wb.Dispose();
            }
        }
    }

    public abstract class WorkbookCodeActivity<TResult> : CodeActivity<TResult>
    {
        protected WorkbookCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<WorkbookCodeActivity<TResult>, WorkbookScope>(Resources.Validation_ScopeErrorFormat(nameof(WorkbookScope))));
        }
        protected abstract override TResult Execute(CodeActivityContext context);
    }
}