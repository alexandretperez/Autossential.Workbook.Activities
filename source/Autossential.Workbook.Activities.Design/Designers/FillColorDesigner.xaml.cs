using Autossential.Shared.Utils;
using Autossential.Workbook.Core.Enums;

namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for FillColorDesigner.xaml
    public partial class FillColorDesigner
    {
        public FillColorDesigner()
        {
            InitializeComponent();

            cbPattern.ItemsSource = EnumUtil.EnumAsDictionary<FillPattern>();
            cbPattern.DisplayMemberPath = "Key";
            cbPattern.SelectedValuePath = "Value";
        }
    }
}
