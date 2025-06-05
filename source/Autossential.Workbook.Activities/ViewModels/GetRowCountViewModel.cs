using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class GetRowCountViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> Range { get; set; }
        public DesignOutArgument<int> Result { get; set; }
        public GetRowCountViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            SheetName.IsRequired = true;
            SheetName.IsPrincipal = true;
            SheetName.DisplayName = "Sheet Name";
            SheetName.Placeholder = "Enter the sheet name";
            SheetName.OrderIndex = orderIndex++;

            Range.IsRequired = false;
            Range.IsPrincipal = false;
            Range.OrderIndex = orderIndex++;

            Result.IsPrincipal = false;
            Result.OrderIndex = orderIndex++;
        }
    }
}
