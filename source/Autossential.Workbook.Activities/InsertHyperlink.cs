using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.IO;

namespace Autossential.Workbook.Activities
{
    public sealed class InsertHyperlink : ScopeAwareCodeActivity<bool, WorkbookScope>
    {
        public InArgument<string> WorkbookPath { get; set; }
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Cell { get; set; }
        public InArgument<string> Link { get; set; }
        public InArgument<string> Label { get; set; }
        public InArgument<string> Tooltip { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null && !UseScope) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
            if (Cell == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Cell)));
            if (Link == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Link)));
        }

        protected override bool Execute(CodeActivityContext context)
        {
            if (UseScope)
                return ExecuteWithScope(context);

            var path = WorkbookPath.Get(context);
            if (!File.Exists(path))
                throw new FileNotFoundException(path);


            IWorkbookAdapter workbook = null;
            try
            {
                workbook = WorkbookAdapterFactory.Create(path);
                return Execute(context, workbook);
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        private bool Execute(ActivityContext context, IWorkbookAdapter adapter)
        {
            var sheetName = SheetName.Get(context);
            var cellAddress = Cell.Get(context);
            var link = Link.Get(context);
            var label = Label.Get(context);
            var tooltip = Tooltip.Get(context);

            adapter.AddHyperlink(sheetName, cellAddress, label, link, tooltip);
            return true;
        }

        protected override bool ExecuteWithScope(ActivityContext context)
        {
            try
            {
                var workbook = WorkbookScope.GetWorkbookInstance(context);
                return Execute(context, workbook);
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
                throw;
            }
        }
    }
}