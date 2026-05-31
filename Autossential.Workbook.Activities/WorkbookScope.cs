using Autossential.Workbook.Activities.Core;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;

namespace Autossential.Workbook.Activities
{
    public sealed class WorkbookScope : NativeActivity
    {
        [Browsable(false)]
        public ActivityAction<IWorkbookProcessor> Body { get; set; }
        protected override bool CanInduceIdle => true;
        public WorkbookScope()
        {
            Body = new ActivityAction<IWorkbookProcessor>
            {
                Argument = new DelegateInArgument<IWorkbookProcessor>(ActivityContextExtensions.WorkbookInstancePropertyName),
                Handler = new Sequence
                {
                    DisplayName = "Do"
                }
            };
        }

        [RequiredArgument]
        public InArgument<string> WorkbookPath { get; set; }
        public OutArgument<string[]> SheetNames { get; set; }
        public InArgument<string> Password { get; set; }

        private IWorkbookProcessor _workbookProcessor;
        protected override void Execute(NativeActivityContext context)
        {
            var path = WorkbookPath.Get(context);
            _workbookProcessor = WorkbookProcessorFactory.OpenOrCreate(path, Password.Get(context));

            SheetNames?.Set(context, _workbookProcessor.GetSheetNames());
            context.ScheduleAction(Body, _workbookProcessor, OnComplete, OnFaulted);
        }

        private void OnFaulted(NativeActivityFaultContext faultContext, Exception propagatedException, ActivityInstance propagatedFrom)
        {
            _workbookProcessor?.Dispose();
        }

        private void OnComplete(NativeActivityContext context, ActivityInstance completedInstance)
        {
            _workbookProcessor?.Dispose();
        }
    }
}
