using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.IO;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetName : ScopeAwareCodeActivity<string, WorkbookScope>
    {
        public InArgument<bool> ContinueOnError { get; set; }
        public InArgument<string> WorkbookPath { get; set; }
        public InArgument<int> SheetIndex { get; set; } = 0;

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (WorkbookPath == null && !UseScope) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
            if (SheetIndex == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetIndex)));
            if (Result == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Result)));
        }

        protected override string Execute(CodeActivityContext context)
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
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
                if (ContinueOnError.Get(context))
                    return null;

                throw;
            }
            finally
            {
                workbook?.Dispose();
            }
        }
        private string Execute(ActivityContext context, IWorkbookAdapter adapter)
        {
            var index = SheetIndex.Get(context);
            return adapter.GetSheetName(index);
        }

        protected override string ExecuteWithScope(ActivityContext context)
        {
            try
            {
                return Execute(context, WorkbookScope.GetWorkbookInstance(context));
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
                if (ContinueOnError.Get(context))
                    return null;

                throw;
            }
        }
    }
}