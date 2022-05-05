using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public class AppendRange : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<DataTable> InputDataTable { get; set; }
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

            var dt = InputDataTable.Get(context);
            if (dt == null)
                throw new ArgumentNullException("DataTable");

            await Task.Run(() => adapter.AppendRange(sheetName, dt));
            return null;
        }
    }
}
