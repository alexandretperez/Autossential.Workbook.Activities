using Autossential.Workbook.Activities.Core;
using System.Activities;
using System.Activities.Statements;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Xml.Linq;

namespace Autossential.Workbook.Activities.Tests
{
    public abstract class BaseTests
    {
        [After(Test)]
        public void Delete()
        {
            if (File.Exists(_filePath))
                File.Delete(_filePath);
        }
        private string? _filePath;

        /// <summary>
        /// Simulates corruption in an Excel file by renaming it to .zip,
        /// extracting and modifying the sharedStrings.xml, replacing '*' with empty values,
        /// reinserting the file into the archive, and restoring the .xlsx extension.
        /// This creates a mismatch between metadata and actual cell content,
        /// making Excel interpret more rows/columns than truly exist.
        /// </summary>
        /// <param name="filePath">The .xlsx file</param>
        protected static void Corrupt(string filePath)
        {
            string tempZip = Path.ChangeExtension(filePath, ".zip");
            string tempSharedStrings = Path.Combine(Path.GetTempPath(), "sharedStrings.xml");

            File.Copy(filePath, tempZip, true);

            using (ZipArchive archive = ZipFile.Open(tempZip, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/sharedStrings.xml");
                entry?.ExtractToFile(tempSharedStrings, true);
            }

            XDocument doc = XDocument.Load(tempSharedStrings);
            foreach (var textNode in doc.Descendants("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            {
                if (textNode.Value.Contains('*'))
                    textNode.Value = textNode.Value.Replace("*", "");
            }
            doc.Save(tempSharedStrings);

            using (ZipArchive archive = ZipFile.Open(tempZip, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/sharedStrings.xml");
                entry?.Delete();
                archive.CreateEntryFromFile(tempSharedStrings, "xl/sharedStrings.xml");
            }

            File.Copy(tempZip, filePath, true);
            File.Delete(tempSharedStrings);
            File.Delete(tempZip);
        }

        protected string NewTempFilePath(string extension)
        {
            _filePath = Path.ChangeExtension(Path.GetTempFileName(), extension);
            return _filePath;
        }

        protected (IWorkbookProcessor processor, string filePath) NewFile(string extension)
        {
            _filePath = NewTempFilePath(extension);
            var processor = WorkbookProcessorFactory.OpenOrCreate(_filePath);
            return (processor, _filePath);
        }

        public static void InvokeWorkbookScopeWith(string workbookPath, Activity activity, string tag = WorkbookScope.TAG)
        {
            var dyn = new DynamicActivity
            {
                Implementation = () => new WorkbookScope
                {
                    WorkbookPath = workbookPath,
                    Body = new ActivityAction<IWorkbookProcessor>
                    {
                        Argument = new DelegateInArgument<IWorkbookProcessor>(tag),
                        Handler = activity
                    }
                }
            };

            WorkflowInvoker.Invoke(dyn);
        }

        public static T InvokeWorkbookScopeWith<T>(string workbookPath, Activity<T> activity, string tag = WorkbookScope.TAG)
        {
            var dyn = new DynamicActivity<T>();
            activity.Result = new OutArgument<T>(env => dyn.Result.Get(env));
            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = workbookPath,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(tag),
                    Handler = activity
                }
            };

            return WorkflowInvoker.Invoke(dyn);
        }

        public static T InvokeWorkbookScopeWith<T>(string workbookPath, Variable[] variables, Activity[] activities, Expression<Func<ActivityContext, T>> outputHandler)
        {
            var dyn = new DynamicActivity<T>();
            var handler = new Sequence();
            foreach (var v in variables)
                handler.Variables.Add(v);

            foreach (var a in activities)
                handler.Activities.Add(a);

            handler.Activities.Add(new Assign<T>
            {
                To = new OutArgument<T>(env => dyn.Result.Get(env)),
                Value = new InArgument<T>(outputHandler)
            });

            dyn.Implementation = () => new WorkbookScope
            {
                WorkbookPath = workbookPath,
                Body = new ActivityAction<IWorkbookProcessor>
                {
                    Argument = new DelegateInArgument<IWorkbookProcessor>(WorkbookScope.TAG),
                    Handler = handler
                }
            };

            return WorkflowInvoker.Invoke(dyn);
        }
    }
}
