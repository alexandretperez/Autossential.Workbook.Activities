using Autossential.Workbook.Activities.Core;
using System.Activities;
using System.ComponentModel;

namespace Autossential.Workbook.Activities.Extensions
{
    public static class ActivityContextExtensions
    {
        extension(ActivityContext context)
        {
            public IWorkbookProcessor GetWorkbookProcessor()
            {
                var properties = context.DataContext.GetProperties();
                var property = properties[WorkbookScope.TAG];
                if (property == null)
                {
                    foreach (PropertyDescriptor prop in properties)
                    {
                        if (prop.PropertyType == typeof(IWorkbookProcessor))
                        {
                            property = prop;
                            break;
                        }
                    }
                }

                return property.GetValue(context.DataContext) as IWorkbookProcessor;
            }
        }
    }
}