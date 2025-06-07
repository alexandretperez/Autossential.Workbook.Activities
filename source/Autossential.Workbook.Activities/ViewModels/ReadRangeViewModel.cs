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
        public DesignProperty<bool> UseColumnDataType { get; set; }
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
            SheetName.DisplayName = "Sheet Name";
            SheetName.Placeholder = "Enter the sheet name";
            SheetName.Category = "Input";
            SheetName.Tooltip = "The sheet name to be read";
            SheetName.OrderIndex = orderIndex++;

            Range.IsRequired = false;
            Range.IsPrincipal = false;
            Range.Category = "Input";
            Range.Placeholder = "Enter the range";
            Range.Tooltip = "Enter the range or leave it blank to read the entire sheet. Ranges can be delimited as A1:E10 or as a start point as just A1.";
            Range.OrderIndex = orderIndex++;

            HasHeaders.IsPrincipal = false;
            HasHeaders.Category = "Options";
            HasHeaders.DisplayName = "Has Headers";
            HasHeaders.Widget = new DefaultWidget { Type = ViewModelWidgetType.NullableBoolean };
            HasHeaders.Tooltip = "Indicates if the data to be read contains headers.";
            HasHeaders.OrderIndex = orderIndex++;

            UseColumnDataType.IsPrincipal = false;
            UseColumnDataType.Category = "Options";
            UseColumnDataType.DisplayName = "Use Column Data Type";
            UseColumnDataType.Widget = new DefaultWidget { Type = ViewModelWidgetType.NullableBoolean };
            UseColumnDataType.Tooltip = "Analyzes the data and if possible, creates the DataTable with the corresponding data type for each column. Otherwise, all columns are set as System.String.";
            UseColumnDataType.OrderIndex = orderIndex++;

            Result.IsPrincipal = false;
            Result.IsRequired = false;
            Result.Category = "Output";
            Result.Tooltip = "The corresponding DataTable";
            Result.OrderIndex = orderIndex++;
        }
    }
}
