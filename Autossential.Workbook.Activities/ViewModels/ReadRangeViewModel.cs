using Autossential.Workbook.Activities.Base;
using System.Activities.DesignViewModels;
using System.Data;

namespace Autossential.Workbook.Activities.ViewModels
{
    internal class ReadRangeViewModel(IDesignServices services) : BaseViewModel(services)
    {
        public DesignInArgument<string> SheetName { get; set; }
        public DesignInArgument<string> Range { get; set; }
        public DesignProperty<bool> HasHeaders { get; set; }
        public DesignInArgument<int> HeaderRows { get; set; }
        public DesignInArgument<int> RowsPerRecord { get; set; }
        public DesignOutArgument<DataTable> Result { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();

            PersistValuesChangedDuringInit();

            var orderIndex = 0;
            SheetName.IsPrincipal = true;
            SheetName.OrderIndex = orderIndex++;

            Range.OrderIndex = orderIndex++;

            HasHeaders.OrderIndex = orderIndex++;
            HeaderRows.OrderIndex = orderIndex++;
            RowsPerRecord.OrderIndex = orderIndex++;
            Result.OrderIndex = orderIndex++;
        }

        protected override void InitializeRules()
        {
            base.InitializeRules();

            Rule(nameof(HasHeaders), () => HeaderRows.IsVisible = HasHeaders.Value, true);
        }
    }
}