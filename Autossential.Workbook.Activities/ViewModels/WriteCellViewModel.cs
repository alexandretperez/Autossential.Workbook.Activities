using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class WriteCellViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> CellAddress { get; set; }
        public DesignInArgument<object> Value { get; set; }
        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            CellAddress.IsPrincipal = true;
            CellAddress.OrderIndex = orderIndex++;

            Value.IsPrincipal = true;
            Value.OrderIndex = orderIndex++;
        }
    }
}
