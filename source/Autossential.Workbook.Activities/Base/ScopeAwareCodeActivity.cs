using Autossential.Shared.Activities.Constraints;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public abstract class ScopeAwareCodeActivity<TScope> : CodeActivity
    {
        protected ScopeAwareCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<ScopeAwareCodeActivity<TScope>, TScope>(Resources.Validation_ScopeErrorFormat(typeof(TScope).Name)));
        }
    }
    public abstract class ScopeAwareCodeActivity<TResult, TScope> : CodeActivity<TResult>
    {
        protected ScopeAwareCodeActivity()
        {
            Constraints.Add(ActivityConstraints.CreateConstraint<ScopeAwareCodeActivity<TScope, TScope>, TScope>(Resources.Validation_ScopeErrorFormat(typeof(TScope).Name)));
        }
    }

    //public abstract class ScopeAwareCodeActivity<T, TScope> : CodeActivity<T>
    //{
    //    private bool _useScope;
    //    private readonly Constraint _constraint;

    //    [Browsable(false)]
    //    public bool UseScope
    //    {
    //        get => _useScope; set
    //        {
    //            _useScope = value;
    //            AddOrRemoveConstraint(value);
    //        }
    //    }

    //    protected ScopeAwareCodeActivity()
    //    {
    //        _constraint = ActivityConstraints.CreateConstraint<ScopeAwareCodeActivity<T, TScope>, TScope>(Resources.Validation_ScopeErrorFormat(typeof(TScope).Name));
    //    }

    //    private void AddOrRemoveConstraint(bool add)
    //    {
    //        if (add) Constraints.Add(_constraint);
    //        else Constraints.Remove(_constraint);
    //    }

    //    protected abstract T ExecuteWithScope(ActivityContext context);
    //}
}