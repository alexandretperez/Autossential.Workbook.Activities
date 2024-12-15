using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetRowCount : WorkbookCodeActivity<int>
    {
        public InArgument<string> SheetName { get; set; }
        public InArgument<string> Range { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
            if (SheetName == null)
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
        }

        protected override int Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context) ?? "A1";
            return context.GetWorkbook().GetRowCount(sheetName, range);
        }
    }
}