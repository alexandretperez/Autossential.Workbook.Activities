using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadColumn : WorkbookCodeActivity<object[]>
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        [RequiredArgument]
        public InArgument<string> StartingCell { get; set; }
        public InArgument<int> Limit { get; set; }
        protected override object[] Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var startingCell = StartingCell.Get(context);
            var limit = Limit.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.ReadColumn_SheetName_DisplayName));

            if (string.IsNullOrEmpty(startingCell))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(Resources.ReadColumn_StartingCell_DisplayName));

            return context.GetWorkbookProcessor().ReadColumn(sheetName, startingCell, limit);
        }
    }
}
