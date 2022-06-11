using Autossential.Shared.Utils;

namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for DrawBorderDesigner.xaml
    public partial class DrawBorderDesigner
    {
        public DrawBorderDesigner()
        {
            InitializeComponent();

            cbBorder.ItemsSource = EnumUtil.EnumAsDictionary<Core.Enums.Border>();
            cbBorder.DisplayMemberPath = "Key";
            cbBorder.SelectedValuePath = "Value";

            cbBorderStyle.ItemsSource = EnumUtil.EnumAsDictionary<Core.Enums.BorderStyle>();
            cbBorderStyle.DisplayMemberPath = "Key";
            cbBorderStyle.SelectedValuePath = "Value";
        }
    }
}
