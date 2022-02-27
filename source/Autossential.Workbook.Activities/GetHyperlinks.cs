using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using System;
using System.Activities;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Activities
{
    public sealed class GetHyperlinks : ScopeAwareCodeActivity<string[], WorkbookScope>
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

        protected override string[] Execute(CodeActivityContext context)
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
                throw;
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        private string[] Execute(ActivityContext context, IWorkbookAdapter adapter)
        {
            var sheetName = SheetName.Get(context);
            var range = CellRange.Get(context);
            return adapter.GetHyperlinks(sheetName, range).ToArray();
        }

        protected override string[] ExecuteWithScope(ActivityContext context)
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