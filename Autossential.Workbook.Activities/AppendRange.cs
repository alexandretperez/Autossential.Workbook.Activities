using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;
using System.Data;

namespace Autossential.Workbook.Activities
{
    public class AppendRange : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<DataTable> DataTable { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.AppendRange_SheetName_DisplayName));

            var data = DataTable.Get(context) ?? throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.AppendRange_DataTable_DisplayName));
            context.GetWorkbookProcessor().AppendRange(sheetName, data);
        }
    }
}
