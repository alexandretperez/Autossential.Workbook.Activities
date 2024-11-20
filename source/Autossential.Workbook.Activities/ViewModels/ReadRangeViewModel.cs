using System.Activities.DesignViewModels;
using System.Activities.ViewModels;
using System.Data;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class ReadRangeViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> Range { get; set; }
        public DesignProperty<bool> HasHeaders { get; set; }
        public DesignOutArgument<DataTable> Result { get; set; }
        public ReadRangeViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            SheetName.IsRequired = true;
            SheetName.IsPrincipal = true;
            SheetName.Category = "Input";
            SheetName.OrderIndex = orderIndex++;

            Range.IsRequired = false;
            Range.IsPrincipal = true;
            Range.Category = "Input";
            Range.OrderIndex = orderIndex++;

            HasHeaders.IsPrincipal = true;
            HasHeaders.Category = "Options";
            HasHeaders.Widget = new DefaultWidget { Type = ViewModelWidgetType.Toggle };
            HasHeaders.OrderIndex = orderIndex++;

            Result.IsPrincipal = false;
            Result.Category = "Output";
            Result.OrderIndex = orderIndex++;
        }
    }
}
