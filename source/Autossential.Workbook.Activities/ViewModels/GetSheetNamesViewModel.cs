using System.Activities.DesignViewModels;
using System.Collections.Generic;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class GetSheetNamesViewModel : DesignPropertiesViewModel
    {
        public DesignOutArgument<string[]> Result { get; set; }
        public GetSheetNamesViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;
            Result.IsPrincipal = true;
            Result.OrderIndex = orderIndex++;
        }
    }


}
