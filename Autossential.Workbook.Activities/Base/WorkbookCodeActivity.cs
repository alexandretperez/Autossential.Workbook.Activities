using Autossential.Workbook.Activities.Helpers;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities.Base
{
    public abstract class WorkbookCodeActivity : CodeActivity
    {
        protected WorkbookCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<WorkbookCodeActivity, WorkbookScope>(ResourcesFn.Common_UsingOutsideOfScopeFormat(Resources.WorkbookScope_DisplayName)));
        }
    }

    public abstract class WorkbookCodeActivity<TResult> : CodeActivity<TResult>
    {
        protected WorkbookCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<WorkbookCodeActivity<TResult>, WorkbookScope>(ResourcesFn.Common_UsingOutsideOfScopeFormat(Resources.WorkbookScope_DisplayName)));
        }
        protected abstract override TResult Execute(CodeActivityContext context);
    }
}
