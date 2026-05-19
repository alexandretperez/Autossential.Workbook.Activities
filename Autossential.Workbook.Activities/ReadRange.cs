using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
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
            var range = Range.Get(context) ?? "A1";
            var headerRows = HeaderRows.Get(context);
            var rowsPerRecord = RowsPerRecord.Get(context);

            return context.GetWorkbookProcessor().ReadRange(sheetName, range, HasHeaders, headerRows, rowsPerRecord);
        }
    }
}
