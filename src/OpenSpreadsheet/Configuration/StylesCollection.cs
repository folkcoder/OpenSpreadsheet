namespace OpenSpreadsheet.Configuration
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using OpenSpreadsheet.Enums;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Stores spreadsheet style definitions.
    /// </summary>
    /// <remarks>For some elements, attributes must be present in a particular order in order to validate against the OpenXml schema.</remarks>
    public class StylesCollection
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StylesCollection"/> class.
        /// </summary>
        public StylesCollection() { }

        /// <summary>
        /// Gets a collection of borders and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.Border, uint> Borders { get; } = new Dictionary<OpenXml.Border, uint>();

        /// <summary>
        /// Gets a collection of cell formats and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.CellFormat, uint> CellFormats { get; } = new Dictionary<OpenXml.CellFormat, uint>();

        /// <summary>
        /// Gets a collection of cell style formats and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.CellFormat, uint> CellStyleFormats { get; } = new Dictionary<OpenXml.CellFormat, uint>();

        /// <summary>
        /// Gets a collection of fills and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.Fill, uint> Fills { get; } = new Dictionary<OpenXml.Fill, uint>();

        /// <summary>
        /// Gets a collection of fonts and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.Font, uint> Fonts { get; } = new Dictionary<OpenXml.Font, uint>();

        /// <summary>
        /// Gets a collection of number formats and their associated stylehseet position index.
        /// </summary>
        public Dictionary<OpenXml.NumberingFormat, uint> NumberingFormats { get; } = new Dictionary<OpenXml.NumberingFormat, uint>();

        /// <summary>
        /// Adds a border to the stylesheet.
        /// </summary>
        /// <param name="placement">The border's cell placement.</param>
        /// <param name="style">The border's style.</param>
        /// <param name="color">The border's color.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddBorder(BorderPlacement placement, OpenXml.BorderStyleValues style, in Color color)
        {
            var border = this.ConstructBorder(placement, style, color);
            return this.ResolveBorderKey(border);
        }

        /// <summary>
        /// Adds a border to the stylesheet.
        /// </summary>
        /// <param name="border">The OpenXml border element to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddBorder(OpenXml.Border border)
        {
            var clonedElement = (OpenXml.Border)border.CloneNode(true);
            return this.ResolveBorderKey(clonedElement);
        }

        /// <summary>
        /// Adds a cell format to the stylesheet.
        /// </summary>
        /// <param name="borderId">The stylesheet position index of the cell format's border property.</param>
        /// <param name="fillId">The stylesheet position index of the cell format's fill property.</param>
        /// <param name="fontId">The stylesheet position index of the cell format's font property.</param>
        /// <param name="cellStyleFormatId">The stylesheet position index of the cell format's cell style format property.</param>
        /// <param name="numberFormatId">The stylesheet position index of the cell format's number format property.</param>
        /// <param name="horizontalAlignment">The cell format's horizontal alignment.</param>
        /// <param name="verticalAlignment">The cell format's vertical alignment.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddCellFormat(
            uint borderId = 0,
            uint fillId = 0,
            uint fontId = 0,
            uint cellStyleFormatId = 0,
            uint numberFormatId = (uint)OpenXmlNumberingFormat.General,
            OpenXml.HorizontalAlignmentValues horizontalAlignment = OpenXml.HorizontalAlignmentValues.General,
            OpenXml.VerticalAlignmentValues verticalAlignment = OpenXml.VerticalAlignmentValues.Center)
        {
            var cellFormat = this.ConstructCellFormat(borderId, fillId, fontId, cellStyleFormatId, numberFormatId, horizontalAlignment, verticalAlignment);
            return this.ResolveCellFormatKey(cellFormat);
        }

        /// <summary>
        /// Adds a cell format to the stylesheet.
        /// </summary>
        /// <param name="cellFormat">The OpenXml cell format element to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddCellFormat(OpenXml.CellFormat cellFormat)
        {
            var clonedElement = (OpenXml.CellFormat)cellFormat.CloneNode(true);
            return this.ResolveCellFormatKey(clonedElement);
        }

        /// <summary>
        /// Adds a cell style format to the stylesheet.
        /// </summary>
        /// <param name="borderId">The stylesheet position index of the cell style format's border property.</param>
        /// <param name="fillId">The stylesheet position index of the cell format's fill property.</param>
        /// <param name="fontId">The stylesheet position index of the cell format's font property.</param>
        /// <param name="cellStyleFormatId">The stylesheet position index of the cell format's cell style format property.</param>
        /// <param name="numberFormatId">The stylesheet position index of the cell format's number format property.</param>
        /// <param name="horizontalAlignment">The cell format's horizontal alignment.</param>
        /// <param name="verticalAlignment">The cell format's vertical alignment.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddCellStyleFormat(
            uint borderId = 0,
            uint fillId = 0,
            uint fontId = 0,
            uint cellStyleFormatId = 0,
            uint numberFormatId = (uint)OpenXmlNumberingFormat.General,
            OpenXml.HorizontalAlignmentValues horizontalAlignment = OpenXml.HorizontalAlignmentValues.General,
            OpenXml.VerticalAlignmentValues verticalAlignment = OpenXml.VerticalAlignmentValues.Center)
        {
            var cellFormat = this.ConstructCellFormat(borderId, fillId, fontId, cellStyleFormatId, numberFormatId, horizontalAlignment, verticalAlignment);
            return this.ResolveCellStyleFormatKey(cellFormat);
        }

        /// <summary>
        /// Adds a cell style format to the stylesheet.
        /// </summary>
        /// <param name="cellStyleFormat">The OpenXml cell style format element to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddCellStyleFormat(OpenXml.CellFormat cellStyleFormat)
        {
            var clonedElement = (OpenXml.CellFormat)cellStyleFormat.CloneNode(true);
            return this.ResolveCellStyleFormatKey(clonedElement);
        }

        /// <summary>
        /// Adds default styles to the styles collection.
        /// </summary>
        public void AddDefaultStyles()
        {
            this.AddBorder(BorderPlacement.None, OpenXml.BorderStyleValues.None, Color.Black);
            this.AddCellFormat();
            this.AddCellStyleFormat();
            this.AddPatternFill(Color.Transparent, OpenXml.PatternValues.None);
            this.AddPatternFill(Color.Transparent, OpenXml.PatternValues.Gray125);
            this.AddFont(new Font(FontFamily.GenericSansSerif, 11, FontStyle.Regular), Color.Black);
        }

        /// <summary>
        /// Adds a fill to the stylesheet.
        /// </summary>
        /// <param name="fill">The OpenXml fill element to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddFill(OpenXml.Fill fill)
        {
            var clonedElement = (OpenXml.Fill)fill.CloneNode(true);
            return this.ResolveFillKey(clonedElement);
        }

        /// <summary>
        /// Adds a font to the styles collection.
        /// </summary>
        /// <param name="font">The font to be added.</param>
        /// <param name="textColor">The font color to be applied.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddFont(Font font, in Color textColor)
        {
            var openXmlFont = new OpenXml.Font();

            if (font.Bold)
            {
                openXmlFont.AppendChild(new OpenXml.Bold());
            }

            if (font.Italic)
            {
                openXmlFont.AppendChild(new OpenXml.Italic());
            }

            if (font.Strikeout)
            {
                openXmlFont.AppendChild(new OpenXml.Strike());
            }

            if (font.Underline)
            {
                openXmlFont.AppendChild(new OpenXml.Underline());
            }

            openXmlFont.AppendChild(new OpenXml.FontSize() { Val = font.Size });

            var colorText = ConvertColorToHex(textColor);
            openXmlFont.AppendChild(ConvertHexColorToOpenXmlColor(colorText));

            openXmlFont.AppendChild(new OpenXml.FontName() { Val = font.Name });

            return this.ResolveFontKey(openXmlFont);
        }

        /// <summary>
        /// Adds a font to the stylesheet.
        /// </summary>
        /// <param name="font">The OpenXml font element to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddFont(OpenXml.Font font)
        {
            var clonedElement = (OpenXml.Font)font.CloneNode(true);
            return this.ResolveFontKey(clonedElement);
        }

        /// <summary>
        /// Adds a numbering format to the stylesheet.
        /// </summary>
        /// <param name="formatCode">The number mask to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddNumberingFormat(string formatCode)
        {
            var numberingFormat = new OpenXml.NumberingFormat() { FormatCode = formatCode, };
            return this.ResolveNumberingFormatKey(numberingFormat);
        }

        /// <summary>
        /// Adds a numbering format to the stylesheet.
        /// </summary>
        /// <param name="numberingFormat">The OpenXml numbering format to be added.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddNumberingFormat(OpenXml.NumberingFormat numberingFormat)
        {
            var clonedElement = (OpenXml.NumberingFormat)numberingFormat.CloneNode(true);
            return this.ResolveNumberingFormatKey(clonedElement);
        }

        /// <summary>
        /// Adds a pattern fill to the stylesheet.
        /// </summary>
        /// <param name="color">The fill's color.</param>
        /// <param name="patternType">The fill's pattern type.</param>
        /// <returns>The stylesheet position index associated with the element.</returns>
        public uint AddPatternFill(in Color color, OpenXml.PatternValues patternType = OpenXml.PatternValues.Solid)
        {
            var patternFill = new OpenXml.PatternFill()
            {
                PatternType = patternType,
            };

            if (color != Color.Transparent)
            {
                var colorText = ConvertColorToHex(color);
                patternFill.ForegroundColor = new OpenXml.ForegroundColor { Rgb = new DocumentFormat.OpenXml.HexBinaryValue() { Value = colorText } };
            }

            var fill = new OpenXml.Fill() { PatternFill = patternFill };
            return this.ResolveFillKey(fill);
        }

        private static string ConvertColorToHex(in Color color) => Color.FromArgb(color.ToArgb()).Name;

        private static OpenXml.Color ConvertHexColorToOpenXmlColor(string hexColor) => new OpenXml.Color() { Rgb = new DocumentFormat.OpenXml.HexBinaryValue() { Value = hexColor } };

        /// <summary>
        /// Creates an <see cref="OpenXml.Border"/> element.
        /// </summary>
        /// <remarks>Appended XML nodes cannot use the same object instance, so that the Color object must be calculated for each border placement.</remarks>
        private OpenXml.Border ConstructBorder(BorderPlacement placement, OpenXml.BorderStyleValues style, in Color color)
        {
            var colorText = ConvertColorToHex(color);
            var border = new OpenXml.Border();

            // TODO: Replace bitwise operation with HasFlag when moved to .NET Core (performance issues fixed).
            if ((placement & BorderPlacement.Left) != 0)
            {
                border.AppendChild(new OpenXml.LeftBorder() { Color = ConvertHexColorToOpenXmlColor(colorText), Style = style });
            }

            if ((placement & BorderPlacement.Right) != 0)
            {
                border.AppendChild(new OpenXml.RightBorder() { Color = ConvertHexColorToOpenXmlColor(colorText), Style = style });
            }

            if ((placement & BorderPlacement.Top) != 0)
            {
                border.AppendChild(new OpenXml.TopBorder() { Color = ConvertHexColorToOpenXmlColor(colorText), Style = style });
            }

            if ((placement & BorderPlacement.Bottom) != 0)
            {
                border.AppendChild(new OpenXml.BottomBorder() { Color = ConvertHexColorToOpenXmlColor(colorText), Style = style });
            }

            if ((placement & BorderPlacement.DiagonalDown) != 0 || (placement & BorderPlacement.DiagonalUp) != 0)
            {
                border.AppendChild(new OpenXml.DiagonalBorder() { Color = ConvertHexColorToOpenXmlColor(colorText), Style = style });
                border.DiagonalDown = (placement & BorderPlacement.DiagonalDown) != 0;
                border.DiagonalUp = (placement & BorderPlacement.DiagonalUp) != 0;
            }

            return border;
        }

        private OpenXml.CellFormat ConstructCellFormat(
            uint borderId = 0,
            uint fillId = 0,
            uint fontId = 0,
            uint cellFormatId = 0,
            uint numberFormatId = (uint)OpenXmlNumberingFormat.General,
            OpenXml.HorizontalAlignmentValues horizontalAlignment = OpenXml.HorizontalAlignmentValues.General,
            OpenXml.VerticalAlignmentValues verticalAlignment = OpenXml.VerticalAlignmentValues.Center)
        {
            bool applyFill = fillId != 0;
            bool applyNumberFormat = numberFormatId != 0;

            var cellFormat = new OpenXml.CellFormat()
            {
                ApplyFill = applyFill,
                ApplyNumberFormat = applyNumberFormat,
                BorderId = borderId,
                FillId = fillId,
                FontId = fontId,
                FormatId = cellFormatId,
                NumberFormatId = numberFormatId,
            };

            cellFormat.AppendChild(new OpenXml.Alignment
            {
                Horizontal = horizontalAlignment,
                Vertical = verticalAlignment,
            });

            return cellFormat;
        }

        private uint ResolveBorderKey(OpenXml.Border border)
        {
            var matchedItem = this.Borders.FirstOrDefault(x => x.Key.OuterXml == border.OuterXml);
            if (matchedItem.Key == null)
            {
                var value = (uint)this.Borders.Count;
                this.Borders.Add(border, value);
                return value;
            }

            return matchedItem.Value;
        }

        private uint ResolveCellFormatKey(OpenXml.CellFormat cellFormat)
        {
            var matchedItem = this.CellFormats.FirstOrDefault(x => x.Key.OuterXml == cellFormat.OuterXml);
            if (matchedItem.Key == null)
            {
                var value = (uint)this.CellFormats.Count;
                this.CellFormats.Add(cellFormat, value);
                return value;
            }

            return matchedItem.Value;
        }

        private uint ResolveCellStyleFormatKey(OpenXml.CellFormat cellStyleFormat)
        {
            var matchedItem = this.CellStyleFormats.FirstOrDefault(x => x.Key.OuterXml == cellStyleFormat.OuterXml);
            if (matchedItem.Key == null)
            {
                var value = (uint)this.CellStyleFormats.Count;
                this.CellStyleFormats.Add(cellStyleFormat, value);
                return value;
            }

            return matchedItem.Value;
        }

        private uint ResolveFillKey(OpenXml.Fill fill)
        {
            var matchedItem = this.Fills.FirstOrDefault(x => x.Key.OuterXml == fill.OuterXml);
            if (matchedItem.Key == null)
            {
                var value = (uint)this.Fills.Count;
                this.Fills.Add(fill, value);
                return value;
            }

            return matchedItem.Value;
        }

        private uint ResolveFontKey(OpenXml.Font font)
        {
            var matchedItem = this.Fonts.FirstOrDefault(x => x.Key.OuterXml == font.OuterXml);
            if (matchedItem.Key == null)
            {
                var value = (uint)this.Fonts.Count;
                this.Fonts.Add(font, value);
                return value;
            }

            return matchedItem.Value;
        }

        private uint ResolveNumberingFormatKey(OpenXml.NumberingFormat numberingFormat)
        {
            const uint minFormatId = 165;

            var matchedItem = this.NumberingFormats.FirstOrDefault(x => x.Key.FormatCode == numberingFormat.FormatCode);
            if (matchedItem.Key == null)
            {
                uint formatId = this.NumberingFormats.Count == 0 ? minFormatId : this.NumberingFormats.Values.Max() + 1;
                numberingFormat.NumberFormatId = formatId;
                this.NumberingFormats.Add(numberingFormat, formatId);

                this.NumberingFormats.TryGetValue(numberingFormat, out uint key);
                return key;
            }

            return matchedItem.Value;
        }
    }
}