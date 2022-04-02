using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public class WriteRange : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Cell { get; set; }
        public InArgument<DataTable> Value { get; set; }
        public bool AddHeaders { get; set; }
        protected override bool CheckWorkbookPath => false;
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var cell = Cell.Get(context);
            var value = Value.Get(context);

            await adapter.WriteRangeAsync(sheetName, cell, value, AddHeaders);
            return null;
        }
    }
}
