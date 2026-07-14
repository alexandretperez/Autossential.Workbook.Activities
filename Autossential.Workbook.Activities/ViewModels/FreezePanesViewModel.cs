using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class FreezePanesViewModel(IDesignServices services) : DesignPropertiesViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<int> ColumnsToFreeze { get; set; }
        public DesignInArgument<int> RowsToFreeze { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            ColumnsToFreeze.OrderIndex = orderIndex++;
            ColumnsToFreeze.IsPrincipal = true;

            RowsToFreeze.OrderIndex = orderIndex++;
            RowsToFreeze.IsPrincipal = true;
        }
    }
}
