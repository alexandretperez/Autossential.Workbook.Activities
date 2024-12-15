using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class SaveWorkbookViewModel : DesignPropertiesViewModel
    {
        public SaveWorkbookViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();

            PersistValuesChangedDuringInit();
        }
    }
}