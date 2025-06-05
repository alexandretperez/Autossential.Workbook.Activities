using Autossential.Shared;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class ActivateSheet : WorkbookCodeActivity
    {
        public InArgument SheetNameOrIndex { get; set; }

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
            if (sheet is string sheetName)
            {
                context.GetWorkbook().ActivateSheet(sheetName);
            }
            else if (sheet is int sheetIndex)
            {
                context.GetWorkbook().ActivateSheet(sheetIndex);
            }
        }
    }
}
