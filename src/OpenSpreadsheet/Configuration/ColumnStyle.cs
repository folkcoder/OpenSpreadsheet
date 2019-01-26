namespace OpenSpreadsheet.Configuration
{
    using System.Drawing;

    using OpenSpreadsheet.Enums;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Encapsulates properties associated with a spreadsheet column.
    /// </summary>
    public class ColumnStyle
    {
        private string customNumberFormat;
        private OpenXmlNumberingFormat numberFormat;

        /// <summary>
        /// Gets or sets the background color.
        /// </summary>
        public virtual Color BackgroundColor { get; set; } = Color.Transparent;

        /// <summary>
        /// Gets or sets the background pattern type.
        /// </summary>
        public virtual OpenXml.PatternValues BackgroundPatternType { get; set; }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        public virtual Color BorderColor { get; set; } = Color.Black;

        /// <summary>
        /// Gets or sets the border placement.
        /// </summary>
        public virtual BorderPlacement BorderPlacement { get; set; }

        /// <summary>
        /// Gets or sets the border style.
        /// </summary>
        public virtual OpenXml.BorderStyleValues BorderStyle { get; set; }

        /// <summary>
        /// Gets or sets a custom numbering format.
        /// </summary>
        public string CustomNumberFormat
        {
            get => this.customNumberFormat;
            set
            {
                this.customNumberFormat = value;
                this.IsNumberFormatSpecified = true;
            }
        }

        /// <summary>
        /// Gets or sets the font.
        /// </summary>
        public virtual Font Font { get; set; } = new Font("Calibri", 11, FontStyle.Regular);

        /// <summary>
        /// Gets or sets the text color.
        /// </summary>
        public virtual Color ForegroundColor { get; set; } = Color.Black;

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public virtual OpenXml.HorizontalAlignmentValues HoizontalAlignment { get; set; }

        /// <summary>
        /// Gets a value indicating whether a number format has been explicitly specified.
        /// </summary>
        public bool IsNumberFormatSpecified { get; private set; }

        /// <summary>
        /// Gets or sets a number format.
        /// </summary>
        public OpenXmlNumberingFormat NumberFormat
        {
            get => this.numberFormat;
            set
            {
                this.numberFormat = value;
                this.IsNumberFormatSpecified = true;
            }
        }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public virtual OpenXml.VerticalAlignmentValues VerticalAlignment { get; set; } = OpenXml.VerticalAlignmentValues.Center;
    }
}