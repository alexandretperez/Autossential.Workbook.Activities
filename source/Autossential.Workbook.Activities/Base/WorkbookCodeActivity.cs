using Autossential.Shared.Activities.Constraints;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public abstract class WorkbookCodeActivity<TResult> : CodeActivity<TResult>
    {
        protected WorkbookCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<WorkbookCodeActivity<TResult>, WorkbookScope>(Resources.Validation_ScopeErrorFormat(nameof(WorkbookScope))));
        }
        protected abstract override TResult Execute(CodeActivityContext context);
    }

    public abstract class WorkbookCodeActivity : CodeActivity
    {
        protected WorkbookCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<WorkbookCodeActivity, WorkbookScope>(Resources.Validation_ScopeErrorFormat(nameof(WorkbookScope))));
        }
        protected abstract override void Execute(CodeActivityContext context);
    }
}