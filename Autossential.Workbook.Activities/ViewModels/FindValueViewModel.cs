using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class FindValueViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> Range { get; set; }
        public DesignInArgument<object> Value { get; set; }
        public DesignOutArgument<string> CellAddress { get; set; }
        public DesignOutArgument<int> RowNumber { get; set; }
        public DesignOutArgument<int> ColumnNumber { get; set; }
        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            SheetName.IsPrincipal = true;
            Value.IsPrincipal = true;
        }
    }
}
