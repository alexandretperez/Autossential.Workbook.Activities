using Autossential.Shared.Activities.Design;
using Autossential.Workbook.Activities.Design.Designers;
using Autossential.Workbook.Activities.Properties;
using System.Activities.Presentation.Metadata;
using System.ComponentModel;

namespace Autossential.Workbook.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public const string MAIN_CATEGORY = "Autossential";
        public const string WORKBOOK = MAIN_CATEGORY + ".Workbook";
        public const string HYPERLINKS = WORKBOOK + ".Hyperlinks";

        public void Register()
        {
            var workbook = new CategoryAttribute(WORKBOOK);
            var hyperlinks = new CategoryAttribute(HYPERLINKS);

            ActivitiesAttributesBuilder.Build(Resources.ResourceManager, builder =>
            {
                builder.SetDefaultCategories(
                    Resources.Input_Category,
                    Resources.Output_Category,
                    Resources.InputOutput_Category,
                    Resources.Options_Category);

                builder.Register<GetSheetNames, GetSheetNamesDesigner>(workbook)
                       .Register<ReadRange, ReadRangeDesigner>(workbook)
                       .Register<WorkbookScope, WorkbookScopeDesigner>(workbook)
                       .Register<WriteRange, WriteRangeDesigner>(workbook)
                       .Register<AppendRange, AppendRangeDesigner>(workbook)
                       .Register<WriteCell, WriteCellDesigner>(workbook)
                       .Register<FillColor, FillColorDesigner>(workbook)
                       .Register<FreezeUnfreezePanes, FreezeUnfreezePanesDesigner>(workbook);

                builder.Register<GetHyperlinks, GetHyperlinksDesigner>(hyperlinks)
                       .Register<InsertHyperlink, InsertHyperlinkDesigner>(hyperlinks)
                       .Register<RemoveHyperlinks, RemoveHyperlinksDesigner>(hyperlinks);

                builder.RegisterToMember(
                    new DescriptionAttribute(Resources.ScopeAwareCodeActivity_UseScope),
                    nameof(ScopeAwareCodeActivity<object, object>.UseScope),
                    typeof(ScopeAwareCodeActivity<,>).GetDerivedTypes());
            });
        }
    }
}