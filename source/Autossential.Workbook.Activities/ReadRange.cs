using Autossential.Workbook.Activities.Extensions;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadRange : WorkbookCodeActivity<DataTable>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public bool UseColumnDataType { get; set; } = true;
        public bool HasHeaders { get; set; } = true;
        protected override DataTable Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context) ?? "A1";
            return context.GetWorkbook().ReadRange(sheetName, range, HasHeaders, UseColumnDataType);
        }
    }
}