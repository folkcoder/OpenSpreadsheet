namespace OpenSpreadsheet.Configuration
{
    using System.Drawing;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Encapsulates properties associated with a worksheet style definition.
    /// </summary>
    public class WorksheetStyle
    {
        /// <summary>
        /// Gets or sets the header background color.
        /// </summary>
        public virtual Color HeaderBackgroundColor { get; set; } = Color.Transparent;

        /// <summary>
        /// Gets or sets the header background pattern type.
        /// </summary>
        public virtual OpenXml.PatternValues HeaderBackgroundPatternType { get; set; }

        /// <summary>
        /// Gets or sets the header font.
        /// </summary>
        public virtual Font HeaderFont { get; set; } = new Font("Calibri", 11, FontStyle.Bold);

        /// <summary>
        /// Gets or sets the header foreground color.
        /// </summary>
        public virtual Color HeaderForegroundColor { get; set; } = Color.Black;

        /// <summary>
        /// Gets or sets the header horizontal alignment.
        /// </summary>
        public virtual OpenXml.HorizontalAlignmentValues HeaderHoizontalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the header row index for writing.
        /// </summary>
        public virtual uint HeaderRowIndex { get; set; } = 1;

        /// <summary>
        /// Gets or set's the header vertical alignment.
        /// </summary>
        public virtual OpenXml.VerticalAlignmentValues HeaderVerticalAlignment { get; set; } = OpenXml.VerticalAlignmentValues.Center;

        /// <summary>
        /// Gets or sets the maximum column width.
        /// </summary>
        public virtual double MaxColumnWidth { get; set; } = 30.0;

        /// <summary>
        /// Gets or sets the minimum column width.
        /// </summary>
        public virtual double MinColumnWidth { get; set; } = 5.0;

        /// <summary>
        /// Gets or sets a value indicating whether the worksheet should automatically apply filters.
        /// </summary>
        public virtual bool ShouldAutoFilter { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the worksheet should adjust column widths to their contents.
        /// </summary>
        /// <remarks>Setting this value to true may have noticeable performance impacts for large data sets.</remarks>
        public virtual bool ShouldAutoFitColumns { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the header row should be frozen.
        /// </summary>
        public virtual bool ShouldFreezeTopRow { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the header row should be written.
        /// </summary>
        public virtual bool ShouldWriteHeaderRow { get; set; } = true;
    }
}