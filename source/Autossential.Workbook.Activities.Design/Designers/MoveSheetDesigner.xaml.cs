using System.Activities;
using System.Activities.Expressions;
using System.Windows;

namespace Autossential.Workbook.Activities.Design.Designers
{
    // Interaction logic for MoveSheetDesigner.xaml
    public partial class MoveSheetDesigner
    {
        public MoveSheetDesigner()
        {
            InitializeComponent();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var sheetName = ModelItem.Properties[nameof(MoveSheet.SheetName)];
            var copySheetName = ModelItem.Properties[nameof(MoveSheet.CopySheetName)];
            if (sheetName.Value != null && copySheetName.Value == null)
            {
                if (sheetName.Value.GetCurrentValue() is InArgument<string> arg)
                {
                    if (arg.Expression is Literal<string>)
                    {
                        copySheetName.SetValue(new InArgument<string>(arg.Expression + "_Copy"));
                    }
                }
            }
        }
    }
}
