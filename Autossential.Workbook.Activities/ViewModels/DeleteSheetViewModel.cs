using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class DeleteSheetViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            SheetName.IsPrincipal = true;
        }
    }
}