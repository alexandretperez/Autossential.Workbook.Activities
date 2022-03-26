using Autossential.Shared.Activities.Base;
using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Activities.Statements;
using System.IO;

namespace Autossential.Workbook.Activities
{
    public class WorkbookScope : ScopeActivity<IWorkbookAdapter>
    {
        public InArgument<string> WorkbookPath { get; set; }
        private IWorkbookAdapter _adapter;

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
        }

        protected override void Execute(NativeActivityContext context)
        {
            var path = WorkbookPath.Get(context);
            _adapter = WorkbookAdapterFactory.Create(path);
            context.ScheduleAction(Body, _adapter, OnComplete, OnFaulted);
        }

        private void OnComplete(NativeActivityContext context, ActivityInstance completedInstance)
        {
            if (_adapter == null)
                return;

            _adapter.SaveAsync();
            _adapter.Dispose();
        }

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            _adapter?.Dispose();
        }

        private const string WorkbookInstanceProperty = "WorkbookInstance";

        protected override void InitializeBody()
        {
            Body = new ActivityAction<IWorkbookAdapter>
            {
                Argument = new DelegateInArgument<IWorkbookAdapter>(WorkbookInstanceProperty),
                Handler = new Sequence()
                {
                    DisplayName = "Do"
                }
            };
        }

        internal static IWorkbookAdapter GetWorkbookInstance(ActivityContext context)
        {
            var workbook = context.DataContext.GetProperties()[WorkbookInstanceProperty]?.GetValue(context.DataContext) as IWorkbookAdapter;
            if (workbook == null)
                throw new InvalidOperationException("Invalid Workbook");

            return workbook;
        }
    }
}