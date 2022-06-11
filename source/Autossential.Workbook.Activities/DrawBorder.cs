using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using Autossential.Workbook.Core.Enums;
using System;
using System.Activities;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class DrawBorder : WorkbookActivity
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public InArgument<Color> Color { get; set; }
        public Border Border { get; set; } = Border.All;
        public BorderStyle BorderStyle { get; set; } = BorderStyle.Thin;
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

            await Task.Run(() => adapter.DrawBorder(sheetName, range, Border, BorderStyle, color));
            return null;
        }
    }
}
