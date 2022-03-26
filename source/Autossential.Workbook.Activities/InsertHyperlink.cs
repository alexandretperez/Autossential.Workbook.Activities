using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class InsertHyperlink : WorkbookActivity<bool>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Cell { get; set; } = "A1";
        public InArgument<string> Link { get; set; }
        public InArgument<string> Label { get; set; }
        public InArgument<string> Tooltip { get; set; }
        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (Cell == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Cell)));
            if (Link == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Link)));
        }
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var cellAddress = Cell.Get(context);
            var link = Link.Get(context);
            var label = Label.Get(context);
            var tooltip = Tooltip.Get(context);

            var result = await adapter.AddHyperLinkAsync(sheetName, cellAddress, label, link, tooltip);
            return ctx => Result.Set(ctx, result);
        }
    }
}