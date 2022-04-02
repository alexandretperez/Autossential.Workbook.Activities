// ################################
// THIS FILE IS AUTO-GENERATE BY T4
// ################################


namespace Autossential.Workbook.Activities.Properties
{
    using System.Resources;


    public partial class Resources
    {
        public static System.Globalization.CultureInfo Culture { get; set; }

        private static object _internalSyncObject;

        /// <summary>
        /// Thread safe lock object used by this class.
        /// </summary>
        public static object InternalSyncObject
        {
            get
            {
                if (_internalSyncObject is null)
                    System.Threading.Interlocked.CompareExchange(ref _internalSyncObject, new object(), null);

                return _internalSyncObject;
            }
        }
        private static ResourceManager _resourceManager;

        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Advanced)]
        public static ResourceManager ResourceManager
        {
            get
            {
                if (_resourceManager is null)
                {
                    System.Threading.Monitor.Enter(InternalSyncObject);

                    try
                    {
                        if (_resourceManager is null)
                            System.Threading.Interlocked.Exchange(ref _resourceManager, new ResourceManager("Autossential.Workbook.Activities.Properties.Resources", typeof(Resources).Assembly));
                    }
                    finally
                    {
                        System.Threading.Monitor.Exit(InternalSyncObject);
                    }
                }
                return _resourceManager;
            }
        }

        // ####### FORMATTERS ###############################################


        /// <summary>
        /// Looks up a localized string similar to 'Please provide a value for {0}.'.
        /// </summary>
        /// <param name="arg0">An object (0) to format.</param>
        /// <returns>A copy of format string in which the format items have been replaced by the String equivalent of the corresponding instances of Object in arguments.</returns>
        public static string Validation_ValueErrorFormat(object arg0)
        {
            return string.Format(Culture, Validation_ValueError, arg0);
        }

        /// <summary>
        /// Looks up a localized string similar to 'Cannot be used outside of {0} activity.'.
        /// </summary>
        /// <param name="arg0">An object (0) to format.</param>
        /// <returns>A copy of format string in which the format items have been replaced by the String equivalent of the corresponding instances of Object in arguments.</returns>
        public static string Validation_ScopeErrorFormat(object arg0)
        {
            return string.Format(Culture, Validation_ScopeError, arg0);
        }

        /// <summary>
        /// Looks up a localized string similar to 'Please provide a value of type {0} for {1}.'.
        /// </summary>
        /// <param name="arg0">An object (0) to format.</param>
        /// <param name="arg1">An object (1) to format.</param>
        /// <returns>A copy of format string in which the format items have been replaced by the String equivalent of the corresponding instances of Object in arguments.</returns>
        public static string Validation_TypeErrorFormat(object arg0, object arg1)
        {
            return string.Format(Culture, Validation_TypeError, arg0, arg1);
        }


        // ####### GETTERS ##################################################


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string WorkbookScope_WorkbookPath_Description => ResourceManager.GetString("WorkbookScope_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Workbook Scope'.
        /// </summary>
        public static string WorkbookScope_DisplayName => ResourceManager.GetString("WorkbookScope_DisplayName", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The cell address to insert the hyperlink. E.g: "A3", "D5", "F9"...'.
        /// </summary>
        public static string InsertHyperlink_Cell_Description => ResourceManager.GetString("InsertHyperlink_Cell_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Please provide a value for {0}.'.
        /// </summary>
        public static string Validation_ValueError => ResourceManager.GetString("Validation_ValueError", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'A zero-based sheet index to search for.'.
        /// </summary>
        public static string GetSheetName_SheetIndex_Description => ResourceManager.GetString("GetSheetName_SheetIndex_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string RemoveHyperlinks_WorkbookPath_Description => ResourceManager.GetString("RemoveHyperlinks_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The hyperlinks found in the specified range.'.
        /// </summary>
        public static string GetHyperlinks_Result_Description => ResourceManager.GetString("GetHyperlinks_Result_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Gets the sheet name from a workbook at the specified index.'.
        /// </summary>
        public static string GetSheetName_Description => ResourceManager.GetString("GetSheetName_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Uses the workbook instance given by Workbook Scope.'.
        /// </summary>
        public static string ScopeAwareCodeActivity_UseScope => ResourceManager.GetString("ScopeAwareCodeActivity_UseScope", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Cannot be used outside of {0} activity.'.
        /// </summary>
        public static string Validation_ScopeError => ResourceManager.GetString("Validation_ScopeError", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Specifies the amount of time in milliseconds to wait for the activity to run before an error is thrown. The default value is 30000 (30s).'.
        /// </summary>
        public static string Common_Timeout => ResourceManager.GetString("Common_Timeout", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The cell range to look for hyperlinks.'.
        /// </summary>
        public static string GetHyperlinks_CellRange_Description => ResourceManager.GetString("GetHyperlinks_CellRange_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Insert Hyperlinks'.
        /// </summary>
        public static string InsertHyperlink_DisplayName => ResourceManager.GetString("InsertHyperlink_DisplayName", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'If set, continue executing the remaining activities even if the current activity has failed.'.
        /// </summary>
        public static string Common_ContinueOnError => ResourceManager.GetString("Common_ContinueOnError", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string InsertHyperlink_WorkbookPath_Description => ResourceManager.GetString("InsertHyperlink_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Removes the hyperlinks from the specified cell range.'.
        /// </summary>
        public static string RemoveHyperlinks_Description => ResourceManager.GetString("RemoveHyperlinks_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The link to navigate to. Can be an address to URL, Email, File or Workbook Reference.'.
        /// </summary>
        public static string InsertHyperlink_Link_Description => ResourceManager.GetString("InsertHyperlink_Link_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Get Sheet Name'.
        /// </summary>
        public static string GetSheetName_DisplayName => ResourceManager.GetString("GetSheetName_DisplayName", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Get Sheet Names'.
        /// </summary>
        public static string GetSheetNames_DisplayName => ResourceManager.GetString("GetSheetNames_DisplayName", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Inserts hyperlinks on the specified cell address.'.
        /// </summary>
        public static string InsertHyperlink_Description => ResourceManager.GetString("InsertHyperlink_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string GetSheetName_WorkbookPath_Description => ResourceManager.GetString("GetSheetName_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Reference'.
        /// </summary>
        public static string InputOutput_Category => ResourceManager.GetString("InputOutput_Category", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Common'.
        /// </summary>
        public static string Common_Category => ResourceManager.GetString("Common_Category", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The sheet name.'.
        /// </summary>
        public static string RemoveHyperlinks_SheetName_Description => ResourceManager.GetString("RemoveHyperlinks_SheetName_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The number of links removed on the specified range.'.
        /// </summary>
        public static string RemoveHyperlinks_Result_Description => ResourceManager.GetString("RemoveHyperlinks_Result_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Output'.
        /// </summary>
        public static string Output_Category => ResourceManager.GetString("Output_Category", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Gets all the sheet names from a workbook.'.
        /// </summary>
        public static string GetSheetNames_Description => ResourceManager.GetString("GetSheetNames_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The label/text to write on the specified cell. E.g: "My WebSite"'.
        /// </summary>
        public static string InsertHyperlink_Label_Description => ResourceManager.GetString("InsertHyperlink_Label_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The cell range to look for hyperlinks and remove them.'.
        /// </summary>
        public static string RemoveHyperlinks_CellRange_Description => ResourceManager.GetString("RemoveHyperlinks_CellRange_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string GetSheetNames_WorkbookPath_Description => ResourceManager.GetString("GetSheetNames_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Shares the same workbook instance across the child activities.'.
        /// </summary>
        public static string WorkbookScope_Description => ResourceManager.GetString("WorkbookScope_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The tooltip of the link. This feature is only works on open xml formats.'.
        /// </summary>
        public static string InsertHyperlink_Tooltip_Description => ResourceManager.GetString("InsertHyperlink_Tooltip_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Get Hyperlinks'.
        /// </summary>
        public static string GetHyperlinks_DisplayName => ResourceManager.GetString("GetHyperlinks_DisplayName", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'A boolean determining whether the operation was successful or not.'.
        /// </summary>
        public static string InsertHyperlink_Result_Description => ResourceManager.GetString("InsertHyperlink_Result_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The sheet names.'.
        /// </summary>
        public static string GetSheetNames_Result_Description => ResourceManager.GetString("GetSheetNames_Result_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Please provide a value of type {0} for {1}.'.
        /// </summary>
        public static string Validation_TypeError => ResourceManager.GetString("Validation_TypeError", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Options'.
        /// </summary>
        public static string Options_Category => ResourceManager.GetString("Options_Category", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The sheet name.'.
        /// </summary>
        public static string GetSheetName_Result_Description => ResourceManager.GetString("GetSheetName_Result_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Input'.
        /// </summary>
        public static string Input_Category => ResourceManager.GetString("Input_Category", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The sheet name.'.
        /// </summary>
        public static string GetHyperlinks_SheetName_Description => ResourceManager.GetString("GetHyperlinks_SheetName_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'The workbook file path.'.
        /// </summary>
        public static string GetHyperlinks_WorkbookPath_Description => ResourceManager.GetString("GetHyperlinks_WorkbookPath_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Gets all the hyperlinks from the specified cell range.'.
        /// </summary>
        public static string GetHyperlinks_Description => ResourceManager.GetString("GetHyperlinks_Description", Culture);


        /// <summary>
        /// Looks up a localized string similar to 'Remove Hyperlinks'.
        /// </summary>
        public static string RemoveHyperlinks_DisplayName => ResourceManager.GetString("RemoveHyperlinks_DisplayName", Culture);



        // ####### RESOURCE NAMES ###########################################

        /// <summary>
        /// Contains all resource names stored in their respective constants.
        /// </summary>
        public static class ResourceNames
        {

            /// <summary>
            /// Stores the resource name 'WorkbookScope_WorkbookPath_Description'.
            /// </summary>
            public const string WorkbookScope_WorkbookPath_Description = "WorkbookScope_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'WorkbookScope_DisplayName'.
            /// </summary>
            public const string WorkbookScope_DisplayName = "WorkbookScope_DisplayName";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Cell_Description'.
            /// </summary>
            public const string InsertHyperlink_Cell_Description = "InsertHyperlink_Cell_Description";


            /// <summary>
            /// Stores the resource name 'Validation_ValueError'.
            /// </summary>
            public const string Validation_ValueError = "Validation_ValueError";


            /// <summary>
            /// Stores the resource name 'GetSheetName_SheetIndex_Description'.
            /// </summary>
            public const string GetSheetName_SheetIndex_Description = "GetSheetName_SheetIndex_Description";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_WorkbookPath_Description'.
            /// </summary>
            public const string RemoveHyperlinks_WorkbookPath_Description = "RemoveHyperlinks_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_Result_Description'.
            /// </summary>
            public const string GetHyperlinks_Result_Description = "GetHyperlinks_Result_Description";


            /// <summary>
            /// Stores the resource name 'GetSheetName_Description'.
            /// </summary>
            public const string GetSheetName_Description = "GetSheetName_Description";


            /// <summary>
            /// Stores the resource name 'ScopeAwareCodeActivity_UseScope'.
            /// </summary>
            public const string ScopeAwareCodeActivity_UseScope = "ScopeAwareCodeActivity_UseScope";


            /// <summary>
            /// Stores the resource name 'Validation_ScopeError'.
            /// </summary>
            public const string Validation_ScopeError = "Validation_ScopeError";


            /// <summary>
            /// Stores the resource name 'Common_Timeout'.
            /// </summary>
            public const string Common_Timeout = "Common_Timeout";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_CellRange_Description'.
            /// </summary>
            public const string GetHyperlinks_CellRange_Description = "GetHyperlinks_CellRange_Description";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_DisplayName'.
            /// </summary>
            public const string InsertHyperlink_DisplayName = "InsertHyperlink_DisplayName";


            /// <summary>
            /// Stores the resource name 'Common_ContinueOnError'.
            /// </summary>
            public const string Common_ContinueOnError = "Common_ContinueOnError";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_WorkbookPath_Description'.
            /// </summary>
            public const string InsertHyperlink_WorkbookPath_Description = "InsertHyperlink_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_Description'.
            /// </summary>
            public const string RemoveHyperlinks_Description = "RemoveHyperlinks_Description";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Link_Description'.
            /// </summary>
            public const string InsertHyperlink_Link_Description = "InsertHyperlink_Link_Description";


            /// <summary>
            /// Stores the resource name 'GetSheetName_DisplayName'.
            /// </summary>
            public const string GetSheetName_DisplayName = "GetSheetName_DisplayName";


            /// <summary>
            /// Stores the resource name 'GetSheetNames_DisplayName'.
            /// </summary>
            public const string GetSheetNames_DisplayName = "GetSheetNames_DisplayName";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Description'.
            /// </summary>
            public const string InsertHyperlink_Description = "InsertHyperlink_Description";


            /// <summary>
            /// Stores the resource name 'GetSheetName_WorkbookPath_Description'.
            /// </summary>
            public const string GetSheetName_WorkbookPath_Description = "GetSheetName_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'InputOutput_Category'.
            /// </summary>
            public const string InputOutput_Category = "InputOutput_Category";


            /// <summary>
            /// Stores the resource name 'Common_Category'.
            /// </summary>
            public const string Common_Category = "Common_Category";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_SheetName_Description'.
            /// </summary>
            public const string RemoveHyperlinks_SheetName_Description = "RemoveHyperlinks_SheetName_Description";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_Result_Description'.
            /// </summary>
            public const string RemoveHyperlinks_Result_Description = "RemoveHyperlinks_Result_Description";


            /// <summary>
            /// Stores the resource name 'Output_Category'.
            /// </summary>
            public const string Output_Category = "Output_Category";


            /// <summary>
            /// Stores the resource name 'GetSheetNames_Description'.
            /// </summary>
            public const string GetSheetNames_Description = "GetSheetNames_Description";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Label_Description'.
            /// </summary>
            public const string InsertHyperlink_Label_Description = "InsertHyperlink_Label_Description";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_CellRange_Description'.
            /// </summary>
            public const string RemoveHyperlinks_CellRange_Description = "RemoveHyperlinks_CellRange_Description";


            /// <summary>
            /// Stores the resource name 'GetSheetNames_WorkbookPath_Description'.
            /// </summary>
            public const string GetSheetNames_WorkbookPath_Description = "GetSheetNames_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'WorkbookScope_Description'.
            /// </summary>
            public const string WorkbookScope_Description = "WorkbookScope_Description";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Tooltip_Description'.
            /// </summary>
            public const string InsertHyperlink_Tooltip_Description = "InsertHyperlink_Tooltip_Description";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_DisplayName'.
            /// </summary>
            public const string GetHyperlinks_DisplayName = "GetHyperlinks_DisplayName";


            /// <summary>
            /// Stores the resource name 'InsertHyperlink_Result_Description'.
            /// </summary>
            public const string InsertHyperlink_Result_Description = "InsertHyperlink_Result_Description";


            /// <summary>
            /// Stores the resource name 'GetSheetNames_Result_Description'.
            /// </summary>
            public const string GetSheetNames_Result_Description = "GetSheetNames_Result_Description";


            /// <summary>
            /// Stores the resource name 'Validation_TypeError'.
            /// </summary>
            public const string Validation_TypeError = "Validation_TypeError";


            /// <summary>
            /// Stores the resource name 'Options_Category'.
            /// </summary>
            public const string Options_Category = "Options_Category";


            /// <summary>
            /// Stores the resource name 'GetSheetName_Result_Description'.
            /// </summary>
            public const string GetSheetName_Result_Description = "GetSheetName_Result_Description";


            /// <summary>
            /// Stores the resource name 'Input_Category'.
            /// </summary>
            public const string Input_Category = "Input_Category";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_SheetName_Description'.
            /// </summary>
            public const string GetHyperlinks_SheetName_Description = "GetHyperlinks_SheetName_Description";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_WorkbookPath_Description'.
            /// </summary>
            public const string GetHyperlinks_WorkbookPath_Description = "GetHyperlinks_WorkbookPath_Description";


            /// <summary>
            /// Stores the resource name 'GetHyperlinks_Description'.
            /// </summary>
            public const string GetHyperlinks_Description = "GetHyperlinks_Description";


            /// <summary>
            /// Stores the resource name 'RemoveHyperlinks_DisplayName'.
            /// </summary>
            public const string RemoveHyperlinks_DisplayName = "RemoveHyperlinks_DisplayName";

        }
    }
}


