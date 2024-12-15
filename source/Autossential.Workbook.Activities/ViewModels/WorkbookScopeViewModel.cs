using Microsoft.Win32;
using System.Activities.DesignViewModels;
using System.Activities.ViewModels;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class WorkbookScopeViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> WorkbookPath { get; set; }
        public DesignProperty<bool> AutoSave { get; set; }

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
            WorkbookPath.DisplayName = "Workbook Path";
            WorkbookPath.Placeholder = "Enter the workbook path";
            WorkbookPath.OrderIndex = orderIndex++;
            WorkbookPath.AddMenuAction(new MenuAction()
            {
                DisplayName = "Browse for file",
                IsVisible = true,
                IsMain = true,
                Handler = _ => Task.Run(() =>
                {
                    var ofd = new OpenFileDialog
                    {
                        Filter = "All Supported Types (*.xlsx;*.xlsm;*.xltx;*.xltm;*.xls)|*.xlsx;*.xlsm;*.xltx;*.xltm;*.xls|Excel Workbook (*.xlsx)|*.xlsx|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Template (*.xltx)|*.xltx|Excel Macro-Enabled Template (*.xltm)|*.xltm|Excel 97-2003 Workbook (*.xls)|*.xls",
                        Multiselect = false
                    };

                    if (ofd.ShowDialog() == true)
                    {
                        WorkbookPath.Value = ofd.FileName;
                    }
                })
            });

            AutoSave.IsPrincipal = false;
            AutoSave.DisplayName = "Auto Save";
            AutoSave.OrderIndex = orderIndex++;
            AutoSave.Category = "Options";
        }
    }
}