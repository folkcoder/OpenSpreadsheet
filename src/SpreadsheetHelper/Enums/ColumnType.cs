namespace SpreadsheetHelper.Enums
{
    /// <summary>
    /// Encapsulates valid spreadsheet column types.
    /// </summary>
    public enum ColumnType
    {
        /// <summary>
        /// No explicit column type defined.
        /// </summary>
        Unset,

        /// <summary>
        /// A column containing boolean values.
        /// </summary>
        Boolean,

        /// <summary>
        /// A column containing datetime values.
        /// </summary>
        Date,

        /// <summary>
        /// A column containing formulas.
        /// </summary>
        Formula,

        /// <summary>
        /// A column containing numeric values.
        /// </summary>
        Number,

        /// <summary>
        /// A column containing rich text values.
        /// </summary>
        RichText,

        /// <summary>
        /// A column containing text values.
        /// </summary>
        Text,
    }
}