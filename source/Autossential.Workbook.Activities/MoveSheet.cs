using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class MoveSheet : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<int> SheetIndex { get; set; } = 0;
        public InArgument<string> CopySheetName { get; set; }
        public bool MakeACopy { get; set; }
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var sheetIndex = SheetIndex.Get(context);
            var copySheetName = CopySheetName.Get(context);

            await Task.Run(() => adapter.MoveSheet(sheetName, sheetIndex, MakeACopy, copySheetName));
            return null;
        }
    }
}
