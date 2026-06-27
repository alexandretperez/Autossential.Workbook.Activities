using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadRange : WorkbookCodeActivity<DataTable>
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public bool HasHeaders { get; set; } = true;
        public InArgument<int> HeaderRows { get; set; } = 1;
        public InArgument<int> RowsPerRecord { get; set; } = 1;

        protected override DataTable Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var headerRows = HeaderRows.Get(context);
            var rowsPerRecord = RowsPerRecord.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.ReadRange_SheetName_DisplayName));

            if (string.IsNullOrEmpty(range))
                range = "A1";

            return context.GetWorkbookProcessor().ReadRange(sheetName, range, HasHeaders, headerRows, rowsPerRecord);
        }
    }
}
