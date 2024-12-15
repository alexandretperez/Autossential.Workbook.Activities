using Autossential.Shared.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.Activities.Statements;

namespace Autossential.Workbook.Activities
{
    public class WorkbookScope : ScopeActivity<IWorkbookProcessor> // This base class exposes an OutArgument named Result
    {
        [RequiredArgument]
        public InArgument<string> WorkbookPath { get; set; }
        public bool AutoSave { get; set; } = true;

        private IWorkbookProcessor _workbook;

        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
        }

        protected override void Execute(NativeActivityContext context)
        {
            var path = WorkbookPath.Get(context);
            _workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            context.ScheduleAction(Body, _workbook, OnComplete, OnFaulted);
        }

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            _workbook?.Dispose();
        }

        private void OnComplete(NativeActivityContext context, ActivityInstance completedInstance)
        {
            if (_workbook == null) return;
            if (AutoSave)
                _workbook.Save();

            _workbook?.Dispose();
        }

        protected override void InitializeBody()
        {
            Body = new ActivityAction<IWorkbookProcessor>
            {
                Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                Handler = new Sequence()
                {
                    DisplayName = "Do"
                }
            };
        }
    }
}