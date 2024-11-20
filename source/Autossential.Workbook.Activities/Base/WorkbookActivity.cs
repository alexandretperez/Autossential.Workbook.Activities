using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public abstract class WorkbookActivity : ScopeAwareCodeActivityAsync<WorkbookScope>
    {
        public InArgument<string> WorkbookPath { get; set; }

        protected virtual bool CheckWorkbookPath => true;

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (WorkbookPath == null && !UseScope) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(WorkbookPath)));
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken token)
        {
            if (UseScope)
                return await ExecuteAsync(context, WorkbookScope.GetWorkbookInstance(context), token);

            var path = WorkbookPath.Get(context);
            if (path == null) throw new ArgumentNullException(nameof(WorkbookPath));
            if (CheckWorkbookPath && !File.Exists(path))
                throw new FileNotFoundException(path);

            using (var adapter = WorkbookFactory.Create(path))
            {
                var result = await ExecuteAsync(context, adapter, token);
                adapter.Save();
                return result;
            }
        }

        public abstract Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token);
    }

    public abstract class WorkbookActivity<T> : WorkbookActivity
    {
        public OutArgument<T> Result { get; set; }
    }
}