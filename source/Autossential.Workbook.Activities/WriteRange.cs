using Autossential.Workbook.Activities.Properties;
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
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> StartingCell { get; set; }
        public InArgument<DataTable> InputDataTable { get; set; }
        public bool AddHeaders { get; set; }
        protected override bool CheckWorkbookPath => false;

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
            if (InputDataTable == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat("DataTable"));
        }
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            var cell = StartingCell.Get(context);
            if (string.IsNullOrEmpty(cell))
                cell = "A1";

            var dt = InputDataTable.Get(context);
            if (dt == null)
                throw new ArgumentNullException("DataTable");

            await Task.Run(() => adapter.WriteRange(sheetName, cell, dt, AddHeaders));
            return null;
        }
    }
}
