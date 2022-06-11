using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class DeleteSheet : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            await Task.Run(() => adapter.DeleteSheet(sheetName));
            return null;
        }
    }
}
