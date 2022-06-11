using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class RenameSheet : WorkbookActivity
    {
        public InArgument<int> SheetIndex { get; set; }
        public InArgument<string> NewSheetName { get; set; }
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetIndex = SheetIndex.Get(context);
            var newName = NewSheetName.Get(context);
            await Task.Run(() => adapter.RenameSheet(sheetIndex, newName));
            return null;
        }
    }
}
