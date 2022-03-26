using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetNames : WorkbookActivity<string[]>
    {
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var names = await adapter.GetSheetNamesAsync();
            return ctx => Result.Set(ctx, names);
        }
    }
}
