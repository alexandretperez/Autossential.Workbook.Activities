using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities.DesignViewModels;
using System.Activities.ViewModels;
using UiPath.Studio.Activities.Api;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class WorkbookScopeViewModel : BaseViewModel
    {
        public DesignInArgument<string> WorkbookPath { get; set; }
        public DesignOutArgument<string[]> SheetNames { get; set; }
        public DesignInArgument<string> Password { get; set; }

        public WorkbookScopeViewModel(IDesignServices services) : base(services)
        {
            var api = GetWorkflowDesignApi();
            if (api.HasFeature(DesignFeatureKeys.ScopedActivities))
                api.ScopedActivitiesService.SetScopedActivities(typeof(WorkbookScope), [
                    typeof(ReadRange),
                    typeof(GetRangeSize),
                    typeof(ReadCell),
                    typeof(ReadRow),
                    typeof(ReadColumn)
                 ]);
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();

            PersistValuesChangedDuringInit();

            int orderIndex = 0;
            WorkbookPath.IsRequired = true;
            WorkbookPath.IsPrincipal = true;
            WorkbookPath.Placeholder = Resources.WorkbookScope_WorkbookPath_Placeholder;
            WorkbookPath.OrderIndex = orderIndex++;

            Password.OrderIndex = orderIndex++;
            SheetNames.OrderIndex = orderIndex++;

#if WINDOWS
            if (IsWidgetSupported(ViewModelWidgetType.ActionButton))
            {
                WorkbookPath.AddFileDialogMenuAction(true, "Workbook Files|*.xlsx;*.xlsm;*.xls|All Files|*.*");
            }
#endif
        }
    }
}