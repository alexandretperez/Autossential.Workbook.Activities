namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for GetHyperlinksDesigner.xaml
    public partial class GetHyperlinksDesigner
    {
        public GetHyperlinksDesigner()
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
