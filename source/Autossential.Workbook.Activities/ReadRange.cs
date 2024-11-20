using Autossential.Workbook.Activities.Extensions;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadRange : WorkbookCodeActivity<DataTable>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public bool HasHeaders { get; set; } = true;
        protected override DataTable Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context) ?? "A1";
            var wb = context.GetWorkbook();
            try
            {
                return wb.ReadRange(sheetName, range, HasHeaders);
            }
            finally
            {
                wb.Dispose();
            }
        }
    }
}