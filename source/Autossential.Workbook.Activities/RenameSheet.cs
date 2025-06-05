using Autossential.Shared;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class RenameSheet : WorkbookCodeActivity
    {
        public InArgument SheetNameOrIndex { get; set; }

        public InArgument<string> NewSheetName { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetNameOrIndex == null)
            {
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetNameOrIndex)));
            }
            else if (SheetNameOrIndex.IsArgumentTypeAnyCompatible<int, string>())
            {
                metadata.AddRuntimeArgument(SheetNameOrIndex, SheetNameOrIndex.ArgumentType, nameof(SheetNameOrIndex), true);
            }
            else
            {
                metadata.AddValidationError(Resources.Validation_TypeErrorFormat("int or string", nameof(SheetNameOrIndex)));
            }
        }

        protected override void Execute(CodeActivityContext context)
        {
            var sheet = SheetNameOrIndex.Get(context);
            var newSheetName = NewSheetName.Get(context);

            if (sheet is string sheetName)
            {
                context.GetWorkbook().RenameSheet(sheetName, newSheetName);
            }
            else if (sheet is int sheetIndex)
            {
                context.GetWorkbook().RenameSheet(sheetIndex, newSheetName);
            }
        }
    }
}
