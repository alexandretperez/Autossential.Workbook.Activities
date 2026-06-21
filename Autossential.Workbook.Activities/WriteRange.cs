using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities
{
    public sealed class WriteRange : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> StartingCell { get; set; }
        public bool AddHeaders { get; set; } = true;
        [RequiredArgument]
        public InArgument<DataTable> DataTable { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var startingCell = StartingCell.Get(context) ?? "A1";
            
            if (startingCell.Length == 0)
                startingCell = "A1";

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.WriteRange_SheetName_DisplayName));

            var data = DataTable.Get(context);
            context.GetWorkbookProcessor().WriteRange(sheetName, data, startingCell, AddHeaders);
        }
    }
}
