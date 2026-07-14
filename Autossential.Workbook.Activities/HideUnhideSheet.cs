using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;
using System.ComponentModel.DataAnnotations;

namespace Autossential.Workbook.Activities
{
    public sealed class HideUnhideSheet : WorkbookCodeActivity
    {
        public enum HideUnhideAction
        {
            Hide,
            Unhide
        }

        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        public HideUnhideAction Action { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException(ResourcesFn.Common_ErrorMsg_ValueNotSuppliedFormat(""));

            if (Action == HideUnhideAction.Hide)
                context.GetWorkbookProcessor().HideSheet(sheetName);
            else
                context.GetWorkbookProcessor().UnhideSheet(sheetName);
        }
    }
}
