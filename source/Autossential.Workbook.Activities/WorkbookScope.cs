using Autossential.Shared.Activities.Base;
using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.Activities.Statements;
using System.IO;

namespace Autossential.Workbook.Activities
{
    public class WorkbookScope : ScopeActivity<IWorkbookAdapter>
    {
        public InArgument<string> WorkbookPath { get; set; }
        private IWorkbookAdapter _workbook;

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
        }

        protected override void Execute(NativeActivityContext context)
        {
            var path = WorkbookPath.Get(context);
            if (!File.Exists(path))
                throw new FileNotFoundException(path);

            _workbook = WorkbookAdapterFactory.Create(path);
            context.ScheduleAction(Body, _workbook, OnComplete, OnFaulted);
        }

        private void OnComplete(NativeActivityContext context, ActivityInstance completedInstance)
        {
            _workbook?.Dispose();
        }

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            _workbook?.Dispose();
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