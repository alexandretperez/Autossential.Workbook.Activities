namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for GetSheetNamesDesigner.xaml
    public partial class GetSheetNamesDesigner
    {
        public GetSheetNamesDesigner()
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
