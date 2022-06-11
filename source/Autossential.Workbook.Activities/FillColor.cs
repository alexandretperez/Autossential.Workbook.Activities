using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Autossential.Shared;
using System.Collections.Generic;
using System.Drawing;
using Autossential.Workbook.Core.Enums;

namespace Autossential.Workbook.Activities
{
    public sealed class FillColor : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public InArgument Color { get; set; }
        public FillPattern Pattern { get; set; }
        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null) metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));

            if (Range == null)
            {
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Range)));
            }

            if (Pattern == FillPattern.None)
                return;

            if (Color == null)
            {
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Color)));
            }
            else if (Color.IsArgumentTypeAnyCompatible<Color, IEnumerable<Color>>())
            {
                metadata.AddRuntimeArgument(Color, Color.ArgumentType, nameof(Color), true);
            }
            else
            {
                metadata.AddValidationError(Resources.Validation_TypeErrorFormat("Color or IEnumerable<Color>", nameof(Color)));
            }
        }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            var range = Range.Get(context);
            var color = Color.Get(context);

            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            if (string.IsNullOrEmpty(range))
                throw new ArgumentException(nameof(Range), Resources.Validation_NullOrEmptyFormat(nameof(Range)));

            if (color == null)
                throw new ArgumentException(nameof(Color));

            var colors = color is Color c ? new[] { c } : (Color[])color;

            await Task.Run(() => adapter.FillColor(sheetName, range, colors, Pattern));
            return null;
        }
    }
}
