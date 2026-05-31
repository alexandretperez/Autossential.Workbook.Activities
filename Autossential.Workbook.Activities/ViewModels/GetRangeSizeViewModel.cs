using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class GetRangeSizeViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> Range { get; set; }
        public DesignOutArgument<int> RowCount { get; set; }
        public DesignOutArgument<int> ColumnCount { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;

            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            Range.OrderIndex = orderIndex++;
            RowCount.OrderIndex = orderIndex++;
            ColumnCount.OrderIndex = orderIndex++;
        }
    }
}
