using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadRange : WorkbookActivity<DataTable>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; } = "A1:A2";
        public bool AddHeaders { get; set; } = true;

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);

            var dt = await adapter.ReadRangeAsync(sheetName, range, AddHeaders);
            return ctx => Result.Set(ctx, dt);
        }

    }
}