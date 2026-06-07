namespace Autossential.Workbook.Activities.Core
{
    internal enum RangeOrigin
    {
        /// <summary>
        /// Empty initialization
        /// </summary>
        Default,
        
        /// <summary>
        /// Start and end inferred
        /// </summary>
        Inferred,
        
        /// <summary>
        /// Only end inferred
        /// </summary>
        InferredEnd,

        /// <summary>
        /// Explicit start and end
        /// </summary>
        Explicit
    }
}