using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class ActivateSheetViewModel : DesignPropertiesViewModel
    {
        public ActivateSheetViewModel(IDesignServices services) : base(services)
        {
        }

        public DesignInArgument SheetNameOrIndex { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            SheetNameOrIndex.IsRequired = true;
            SheetNameOrIndex.IsPrincipal = true;
            SheetNameOrIndex.OrderIndex = orderIndex++;
            SheetNameOrIndex.Placeholder = "Enter the sheet name or zero-based index";
            SheetNameOrIndex.DisplayName = "Sheet Name or Index";
            SheetNameOrIndex.Tooltip = "Enter the sheet name or zero-based index";
        }
    }
}
