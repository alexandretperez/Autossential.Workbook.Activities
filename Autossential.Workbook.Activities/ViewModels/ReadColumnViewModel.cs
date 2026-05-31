using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class ReadColumnViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> StartingCell { get; set; }
        public DesignInArgument<int> Limit { get; set; }
        public DesignOutArgument<object[]> Result { get; set; }
        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            StartingCell.OrderIndex++;
            StartingCell.IsPrincipal = true;
            Limit.OrderIndex++;
            Result.OrderIndex++;
        }
    }
}
