using Autossential.Workbook.Core.Adapters;
using Autossential.Workbook.Core.Enums;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class DrawBorder : WorkbookActivity
    {
        public Border Border { get; set; }
        public BorderStyle BorderStyle { get; set; }
        public override Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            throw new NotImplementedException();
        }
    }
}
