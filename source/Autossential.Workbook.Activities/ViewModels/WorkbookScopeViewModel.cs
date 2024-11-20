using Microsoft.Win32;
using System.Activities.DesignViewModels;
using System.Activities.ViewModels;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class WorkbookScopeViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> WorkbookPath { get; set; }
        //public DesignInArgument<bool> CreateIfNotExists { get; set; }

        public WorkbookScopeViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();

            PersistValuesChangedDuringInit();

            int orderIndex = 0;
            WorkbookPath.IsRequired = true;
            WorkbookPath.IsPrincipal = true;
            WorkbookPath.Widget = new DefaultWidget { Type = ViewModelWidgetType.Input };
            WorkbookPath.OrderIndex = orderIndex++;
        }
    }
}