namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for RemoveHyperlinksDesigner.xaml
    public partial class RemoveHyperlinksDesigner
    {
        public RemoveHyperlinksDesigner()
        {
            InitializeComponent();
        }
        protected override void OnModelItemChanged(object newItem)
        {
            base.OnModelItemChanged(newItem);
            WorkbookScopeControl.ToggleScope(ModelItem);
        }
    }
}
