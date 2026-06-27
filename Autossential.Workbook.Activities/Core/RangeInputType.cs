namespace Autossential.Workbook.Activities.Core
{
    internal enum RangeInputType
    {
        /// <summary>
        /// Empty initialization
        /// </summary>
        None,

        /// <summary>
        /// Only start column is set
        /// </summary>
        A,

        /// <summary>
        /// Start column and row is set
        /// </summary>
        A1,

        /// <summary>
        /// Start column and End column is set
        /// </summary>
        AB,

        /// <summary>
        /// Start column and row is set, and End column is set
        /// </summary>
        A1B,

        /// <summary>
        /// Start column is set, and End column and row is set
        /// </summary>
        AB1,

        /// <summary>
        /// Start column and row is set, and End column and row is set,
        /// </summary>
        A1B1
    }
}