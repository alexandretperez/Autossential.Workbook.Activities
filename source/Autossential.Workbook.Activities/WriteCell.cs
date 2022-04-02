using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public class WriteCell : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Cell { get; set; }
        public InArgument<object> Value { get; set; }
        protected override bool CheckWorkbookPath => false;
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var cell = Cell.Get(context);
            var value = Value.Get(context);

            await adapter.WriteCellAsync(sheetName, cell, value);
            return null;
        }
    }
}
