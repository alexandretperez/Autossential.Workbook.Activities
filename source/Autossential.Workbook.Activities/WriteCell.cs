using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public class WriteCell : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Cell { get; set; } = "A1";
        public InArgument<object> Value { get; set; }
        protected override bool CheckWorkbookPath => false;
        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
            if (Cell == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Cell)));
        }
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            var cell = Cell.Get(context);
            if (string.IsNullOrEmpty(cell))
                throw new ArgumentException(nameof(Cell), Resources.Validation_NullOrEmptyFormat(nameof(Cell)));

            var value = Value.Get(context);
            await Task.Run(() => adapter.WriteCell(sheetName, cell, value));
            return null;
        }
    }
}
