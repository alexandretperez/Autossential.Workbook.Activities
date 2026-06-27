using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class RenameSheetViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> NewSheetName { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            NewSheetName.IsPrincipal = true;
            NewSheetName.OrderIndex = orderIndex++;
        }
    }
}