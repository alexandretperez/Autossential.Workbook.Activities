using Autossential.Shared;
using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class ActivateSheet : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null)
            {
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
                return;
            }
        }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheet = SheetName.Get(context);
            await Task.Run(() => adapter.ActivateSheet(sheet));
            return null;
        }
    }
}
