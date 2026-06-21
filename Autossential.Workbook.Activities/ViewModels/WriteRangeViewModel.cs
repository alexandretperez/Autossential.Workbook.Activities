using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;
using System.Activities.ViewModels;
using System.Data;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class WriteRangeViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> StartingCell { get; set; }
        public DesignProperty<bool> AddHeaders { get; set; }
        public DesignInArgument<DataTable> DataTable { get; set; }
        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            StartingCell.OrderIndex = orderIndex++;

            DataTable.IsPrincipal = true;
            DataTable.OrderIndex = orderIndex++;

            AddHeaders.OrderIndex = orderIndex++;
            AddWidget(AddHeaders, ViewModelWidgetType.Toggle);
        }
    }
}
