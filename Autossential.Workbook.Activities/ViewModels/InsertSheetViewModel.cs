using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class InsertSheetViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<int?> Position { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;
            Position.OrderIndex = orderIndex++;
        }
    }
}