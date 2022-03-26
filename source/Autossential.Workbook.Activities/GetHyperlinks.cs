using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class GetHyperlinks : WorkbookActivity<string[]>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; } = "A1:A2";

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (SheetName == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
        }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var hyperlinks = await adapter.GetHyperlinksAsync(sheetName, range);
            return ctx => Result.Set(ctx, hyperlinks);
        }
    }
}