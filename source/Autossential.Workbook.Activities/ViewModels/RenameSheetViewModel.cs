using System;
using System.Activities.DesignViewModels;
using System.Activities.ViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class RenameSheetViewModel : DesignPropertiesViewModel
    {
        public RenameSheetViewModel(IDesignServices services) : base(services)
        {
        }

        public DesignInArgument SheetNameOrIndex { get; set; }
        public DesignInArgument<string> NewSheetName { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            SheetNameOrIndex.IsRequired = true;
            SheetNameOrIndex.IsPrincipal = true;
            SheetNameOrIndex.OrderIndex = orderIndex++;
            SheetNameOrIndex.Placeholder = "Enter the sheet name or zero-based index";
            SheetNameOrIndex.DisplayName = "Sheet Name or Index";
            SheetNameOrIndex.Tooltip = "Enter the sheet name or zero-based index";

            NewSheetName.IsRequired = true;
            NewSheetName.IsPrincipal = true;
            NewSheetName.OrderIndex = orderIndex++;
            NewSheetName.Placeholder = "Enter the new sheet name";
            NewSheetName.DisplayName = "New Sheet Name";
            NewSheetName.Tooltip = "Enter the new sheet name";
        }
    }

    /*public class RenameSheetViewModel : DesignPropertiesViewModel
    {
        public DesignInArgument<string> ToSheetName { get; set; }
        public DesignInArgument<int> SheetIndex { get; set; }
        public DesignInArgument<string> FromSheetName { get; set; }
        public DesignProperty<bool> FindByIndex { get; set; }

        public RenameSheetViewModel(IDesignServices services) : base(services)
        {
        }

        protected override void ManualRegisterDependencies()
        {
            base.ManualRegisterDependencies();
            RegisterDependency(FindByIndex, nameof(FindByIndex.Value), nameof(FindByIndex));
        }

        protected override void InitializeRules()
        {
            base.InitializeRules();

            Rule(nameof(FindByIndex), new Action(FindByIndex_Changed));
        }

        private void FindByIndex_Changed()
        {
            var value = FindByIndex.Value;

            SheetIndex.IsVisible = value;
            SheetIndex.IsRequired = value;

            FromSheetName.IsVisible = !value;
            FromSheetName.IsRequired = !value;
        }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            FindByIndex.IsPrincipal = true;
            FindByIndex.IsRequired = true;
            FindByIndex.DisplayName = "Find by Index";
            FindByIndex.Widget = new DefaultWidget { Type = ViewModelWidgetType.Toggle };
            FindByIndex.OrderIndex = orderIndex++;

            SheetIndex.IsRequired = true;
            SheetIndex.IsPrincipal = true;
            SheetIndex.DisplayName = "Sheet Index";
            SheetIndex.Placeholder = "Enter a zero-based index";
            SheetIndex.OrderIndex = orderIndex++;

            FromSheetName.IsRequired = true;
            FromSheetName.IsPrincipal = true;
            FromSheetName.DisplayName = "From Sheet Name";
            FromSheetName.Placeholder = "Enter a sheet name";
            FromSheetName.OrderIndex = orderIndex++;

            ToSheetName.IsRequired = true;
            ToSheetName.IsPrincipal = true;
            ToSheetName.DisplayName = "To Sheet Name";
            ToSheetName.Placeholder = "Enter a sheet name";
            ToSheetName.OrderIndex = orderIndex++;
        }
    }*/
}
