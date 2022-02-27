using System.Activities.Presentation.Model;

namespace Autossential.Workbook.Activities.Design.Controls
{
    // Interaction logic for WorkbookScopeControl.xaml
    public partial class WorkbookScopeControl
    {
        public WorkbookScopeControl()
        {
            InitializeComponent();
        }

        internal void ToggleScope(ModelItem modelItem)
        {
            var parent = modelItem.Parent;
            var scope = false;
            while (parent != null)
            {
                if (parent.ItemType == typeof(WorkbookScope))
                {
                    scope = true;
                    break;
                }

                parent = parent.Parent;
            }

            modelItem.Properties[nameof(ScopeAwareCodeActivity<object, object>.UseScope)].SetValue(scope);
        }
    }
}
