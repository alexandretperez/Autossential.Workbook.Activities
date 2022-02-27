using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetNames : ScopeAwareCodeActivity<string[], WorkbookScope>
    {
        public InArgument<string> WorkbookPath { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null && !UseScope) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
        }

        protected override string[] Execute(CodeActivityContext context)
        {
            if (UseScope)
                return ExecuteWithScope(context);

            IWorkbookAdapter workbook = null;

            try
            {
                var filePath = WorkbookPath.Get(context);
                workbook = WorkbookAdapterFactory.Create(filePath);
                return workbook.GetSheetNames();
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        protected override string[] ExecuteWithScope(ActivityContext context)
        {
            try
            {
                var workbook = WorkbookScope.GetWorkbookInstance(context);
                return workbook.GetSheetNames();
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
                throw;
            }
        }
    }
}
