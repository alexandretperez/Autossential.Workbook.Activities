using Autossential.Shared.Activities.Design.Controls;
using System.Activities.Presentation.Model;
using System.Windows;

namespace Autossential.Workbook.Activities.Design.Controls
{
    // Interaction logic for WorkbookScopeControl.xaml
    public partial class WorkbookScopeControl
    {
        public static readonly DependencyProperty ModelItemProperty = DependencyProperty.Register("ModelItem", typeof(ModelItem), typeof(WorkbookScopeControl), new PropertyMetadata(default(ModelItem)));
        public static readonly DependencyProperty UseScopeProperty = DependencyProperty.Register("UseScope", typeof(bool), typeof(WorkbookScopeControl), new PropertyMetadata(false));

        public WorkbookScopeControl()
        {
            InitializeComponent();
        }

        private void FilePickerControl_Loaded(object sender, RoutedEventArgs e)
        {
            ((FilePickerControl)sender).Filter =
                "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
                + "|Excel Workbook (*.xlsx)|*.xlsx"
                + "|Excel 97-2003 (*.xls)|*.xls";
        }

        public ModelItem ModelItem
        {
            get { return (ModelItem)GetValue(ModelItemProperty); }
            set { SetValue(ModelItemProperty, value); }
        }

        public bool UseScope
        {
            get { return (bool)GetValue(UseScopeProperty); }
            set { SetValue(UseScopeProperty, value); }
        }

        private void root_Loaded(object sender, RoutedEventArgs e)
        {
            if (ModelItem == null)
                return;

            var parent = ModelItem.Parent;
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

            ModelItem.Properties[nameof(ScopeAwareCodeActivity<object, object>.UseScope)].SetValue(scope);
            UseScope = scope;
        }
    }
}