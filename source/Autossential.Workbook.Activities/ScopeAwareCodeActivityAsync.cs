using Autossential.Shared.Activities.Base;
using Autossential.Shared.Activities.Constraints;
using Autossential.Workbook.Activities.Properties;
using System.Activities.Validation;
using System.ComponentModel;

namespace Autossential.Workbook.Activities
{
    public abstract class ScopeAwareCodeActivityAsync<TResult, TScope> : AsyncTaskCodeActivity
    {
        private bool _useScope;
        private readonly Constraint _constraint = null;

        [Browsable(false)]
        public bool UseScope
        {
            get => _useScope; set
            {
                _useScope = value;
                AddOrRemoveConstraint(value);
            }
        }

        public ScopeAwareCodeActivityAsync()
        {
            _constraint = ActivityConstraints.CreateConstraint<ScopeAwareCodeActivityAsync<TResult, TScope>, TScope>(Resources.Validation_ScopeErrorFormat(typeof(TScope).Name));
        }

        private void AddOrRemoveConstraint(bool add)
        {
            if (add)
                Constraints.Add(_constraint);
            else
                Constraints.Remove(_constraint);
        }
    }
}