using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.IO;

namespace Autossential.Workbook.Activities
{

    public sealed class RemoveHyperlinks : ScopeAwareCodeActivity<int, WorkbookScope>
    {
        public InArgument<string> WorkbookPath { get; set; }
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> CellRange { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null && !UseScope) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
            if (SheetName == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
        }

        protected override int Execute(CodeActivityContext context)
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
                return -1;
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        private int Execute(ActivityContext context, IWorkbookAdapter adapter)
        {
            var sheetName = SheetName.Get(context);
            var range = CellRange.Get(context);
            return adapter.RemoveHyperlink(sheetName, range);
        }

        protected override int ExecuteWithScope(ActivityContext context)
        {
            try
            {
                return Execute(context, WorkbookScope.GetWorkbookInstance(context));
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
                throw;
            }
        }
    }
}