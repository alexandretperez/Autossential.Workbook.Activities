namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for GetSheetNameDesigner.xaml
    public partial class GetSheetNameDesigner
    {
        public GetSheetNameDesigner()
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
