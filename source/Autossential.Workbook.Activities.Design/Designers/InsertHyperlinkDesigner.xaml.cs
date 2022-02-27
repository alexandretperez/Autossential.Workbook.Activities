namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for InsertHyperlinkDesigner.xaml
    public partial class InsertHyperlinkDesigner
    {
        public InsertHyperlinkDesigner()
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
