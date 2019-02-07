namespace OpenSpreadsheet.Enums
{
    /// <summary>
    /// Encapsulates valid builtin cell numbering formats.
    /// </summary>
    public enum OpenXmlNumberingFormat : uint
    {
        /// <summary>
        /// A general format that will make a best guess at displaying data.
        /// </summary>
        General = 0,

        /// <summary>
        /// A general format for numeric data (mask "0").
        /// </summary>
        Number = 1,

        /// <summary>
        /// A general format for decimal data (mask "0.00").
        /// </summary>
        Decimal = 2,

        /// <summary>
        /// A numeric format that uses comma separators (mask "#,##0").
        /// </summary>
        NumberWithCommas = 3,

        /// <summary>
        /// A numeric format that rounds decimal points to two places (mask "#,##0.00").
        /// </summary>
        Currency = 4,

        /// <summary>
        /// A numeric format that multiplies the value by 100 and displays the result with a percentage sign, truncating any decimal amounts (mask "0%").
        /// </summary>
        Percentage = 9,

        /// <summary>
        /// A numeric format that multiplies the value by 100 and displays the result with a percentage sign, displaying any decimal amounts (mask "0.00%").
        /// </summary>
        PercentageDecimal = 10,

        /// <summary>
        /// A numeric format that uses exponential notation (mask "0.00E+00").
        /// </summary>
        Scientific = 11,

        /// <summary>
        /// A numeric format that rounds to the nearest single-digit fraction value (mask "# ?/?").
        /// </summary>
        FractionOneDigit = 12,

        /// <summary>
        /// A numeric format that rounds to the nearest two-digit fraction value (mask "# ??/??").
        /// </summary>        
        FractionTwoDigits = 13,

        /// <summary>
        /// A datetime format that displays a date in mm-dd-yy format (e.g., 01-09-19).
        /// </summary>
        /// <remarks>Excel implements this numbering format as m/dd/yyyy rather than the format defined in the standard.</remarks>
        DateMonthDayYear = 14,

        /// <summary>
        /// A datetime format that displays a date in d-mmm-yy format (e.g., 9-Jan-19).
        /// </summary>
        DateDayMonthYear = 15,

        /// <summary>
        /// A datetime format that displays a date in d-mmm format (e.g., 9-Jan).
        /// </summary>
        DateDayMonth = 16,

        /// <summary>
        /// A datetime format that displays a date in mmm-yy format (e.g., Jan-19).
        /// </summary>
        DateMonthYear = 17,

        /// <summary>
        /// A datetime format that displays a time in h:m AM/PM format, using a twelve-hour clock (e.g., 5:54 PM).
        /// </summary>
        TimestampHourMinute12 = 18,

        /// <summary>
        /// A datetime format that displays a time in h:m:s AM/PM format, using a twelve-hour clock (e.g., 5:54:35 PM).
        /// </summary>
        TimestampHouMinuteSecond12 = 19,

        /// <summary>
        /// A datetime format that displays a time in h:m format, using a twenty-four-hour clock (e.g., 17:54).
        /// </summary>
        TimestampHourMinute24 = 20,

        /// <summary>
        /// A datetime format that displays a time in h:m:s format, using a twenty-four-hour clock (e.g., 17:54:35).
        /// </summary>
        TimestampHouMinuteSecond24 = 21,

        /// <summary>
        /// A datetime format that displays a datetime in m/d/yyyy h:m format (e.g., 1/9/2019 17:54).
        /// </summary>
        DatestampWithTime = 22,

        /// <summary>
        /// A numeric format that displays negative numbers surrounded by parentheses (mask "#,##0 ;(#,##0)").
        /// </summary>
        NumberWithNegativeInParens = 37,

        /// <summary>
        /// A numeric format that displays negative numbers in red text, surrounded by parentheses (mask "#,##0 ;[Red](#,##0)").
        /// </summary>
        NumberWithNegativeInRedParens = 38,

        /// <summary>
        /// A numeric format that displays negative numbers surrounded by parentheses (mask "#,##0.00;(#,##0.00)").
        /// </summary>
        DecimalWithNegativeInParens = 39,

        /// <summary>
        /// A decimal format that displays negative numbers in red text, surrounded by parentheses (mask "#,##0.00;[Red](#,##0.00)").
        /// </summary>
        DecimalWithNegativeInRedParens = 40,

        /// <summary>
        /// A decimal format that displays currency amounts (mask "'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';").
        /// </summary>
        /// <remarks>This numbering format is not part of the OpenXml standard and may not be supported by all applications.</remarks>
        Accounting = 44,

        /// <summary>
        /// A timestamp format that displays a time in m:ss format (e.g., 1:15).
        /// </summary>
        TimestampMinuteSecond = 45,

        /// <summary>
        /// A timestamp format that displays a time in [h]:m:ss format, displaying the hour position only if applicable (e.g., 7:1:15).
        /// </summary>
        TimestampConditionalHourMinuteSecond = 46,

        /// <summary>
        /// A format for text.
        /// </summary>
        Text = 49,
    }
}