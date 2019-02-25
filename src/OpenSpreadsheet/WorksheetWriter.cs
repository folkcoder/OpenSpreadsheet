namespace OpenSpreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Globalization;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;
    using OpenSpreadsheet.Extensions;

    /// <summary>
    /// Writes data to a worksheet.
    /// </summary>
    /// <typeparam name="TClass">The class type to be written.</typeparam>
    /// <typeparam name="TClassMap">A map defining how to write record data to the worksheet.</typeparam>
    public class WorksheetWriter<TClass, TClassMap> : IDisposable
        where TClass : class
        where TClassMap : ClassMap<TClass>
    {
        private const string cellTypeAttribute = "t";
        private const uint headerStyleKey = 0;
        private const string rowIndexAttribute = "r";
        private const string styleIndexAttribute = "s";

        private static readonly Dictionary<CellValues, OpenXmlAttribute> cellValueAttributes = new Dictionary<CellValues, OpenXmlAttribute>()
        {
            { CellValues.Boolean, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.Boolean)) },
            { CellValues.Date, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.Number)) },
            { CellValues.InlineString, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.InlineString)) },
            { CellValues.Number, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.Number)) },
            { CellValues.SharedString, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.SharedString)) },
            { CellValues.String, new OpenXmlAttribute(cellTypeAttribute, null, new EnumValue<CellValues>(CellValues.String)) },
        };

        private readonly Dictionary<uint, string> columnCellReferences = new Dictionary<uint, string>();

        /// <summary>
        /// Holds relationship between column indexes (key) to its cell format id in the spreadsheet stylesheet (value). Header cell style has a key of 0.
        /// </summary>
        private readonly Dictionary<uint, uint> columnStyles = new Dictionary<uint, uint>();

        /// <summary>
        /// Holds relationshop between column indexes (key) to its associated CellValues type.
        /// </summary>
        private readonly Dictionary<uint, CellValues> columnTypes = new Dictionary<uint, CellValues>();

        private readonly Dictionary<uint, int> columnWidths = new Dictionary<uint, int>();
        private readonly List<PropertyMap<TClass>> orderedPropertyMaps;
        private readonly BidirectionalDictionary<string, string> sharedStrings;
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly StylesCollection spreadsheetStyles;
        private readonly string worksheetName;
        private readonly int worksheetPositionIndex;
        private readonly WorksheetStyle worksheetStyle;
        private readonly OpenXmlWriter writer;

        private static OpenXmlAttribute cellReferenceAttribute = new OpenXmlAttribute(rowIndexAttribute, null, null);
        private uint currentRowIndex;
        private bool isDisposed = false;
        private WorksheetPart worksheetPart;

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetWriter{TClass, TClassMap}"/> class.
        /// </summary>
        /// <param name="worksheetName">The worksheet's name.</param>
        /// <param name="worksheetPositionIndex">The position where the worksheet should be inserted in the workbook.</param>
        /// <param name="spreadsheetDocument">A reference to the worksheet's parent <see cref="spreadsheetDocument"/>.</param>
        /// <param name="worksheetStyle">The workseet's style definitions.</param>
        /// <param name="sharedStrings">The shared strings collection shared by all worksheets.</param>
        /// <param name="spreadsheetStyles">The spreadsheet stylesheet shared by all worksheets.</param>
        public WorksheetWriter(string worksheetName, int worksheetPositionIndex, SpreadsheetDocument spreadsheetDocument, WorksheetStyle worksheetStyle, BidirectionalDictionary<string, string> sharedStrings, StylesCollection spreadsheetStyles)
        {
            this.sharedStrings = sharedStrings;
            this.spreadsheetDocument = spreadsheetDocument;
            this.spreadsheetStyles = spreadsheetStyles;
            this.worksheetName = worksheetName;
            this.worksheetPositionIndex = worksheetPositionIndex;
            this.worksheetStyle = worksheetStyle;

            // setup worksheet
            this.currentRowIndex = this.worksheetStyle.HeaderRowIndex;
            this.worksheetPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // property maps
            this.orderedPropertyMaps = this.CreateOrderedPropertyMaps().ToList();

            // worksheet setup
            this.CacheWorksheetData();
            this.writer = OpenXmlWriter.Create(this.worksheetPart, System.Text.Encoding.UTF8);
            this.WriteInitialWorksheetElements();
        }

        /// <summary>
        /// Skip a single row.
        /// </summary>
        public void SkipRow() => this.currentRowIndex++;

        /// <summary>
        /// Skip one or more rows.
        /// </summary>
        /// <param name="count">The number of rows to skip.</param>
        public void SkipRows(uint count) => this.currentRowIndex += count;

        /// <summary>
        /// Writes the worksheet header at the current position.
        /// </summary>
        public void WriteHeader()
        {
            this.writer.WriteStartElement(new Row(), ConstructRowAttributes(this.currentRowIndex));
            foreach (var propertyMap in this.orderedPropertyMaps)
            {
                var headerName = ResolveHeader(propertyMap);
                var cellReference = this.ConstructExcelCellReference(this.currentRowIndex, propertyMap.PropertyData.IndexWrite);
                var headerAttributes = new List<OpenXmlAttribute>() { new OpenXmlAttribute(styleIndexAttribute, null, this.columnStyles[headerStyleKey].ToString()) };
                this.WriteCellValue(headerName, cellReference, CellValues.SharedString, headerAttributes);

                if (this.worksheetStyle.ShouldAutoFitColumns)
                {
                    this.AddColumnWidth(propertyMap.PropertyData.IndexWrite, headerName.Length);
                }
            }

            this.currentRowIndex++;
            this.writer.WriteEndElement(); // row
        }

        /// <summary>
        /// Writes the provided record at the current position.
        /// </summary>
        /// <param name="record">The record to be written.</param>
        public void WriteRecord(TClass record)
        {
            this.writer.WriteStartElement(new Row(), ConstructRowAttributes(this.currentRowIndex));
            foreach (var propertyMap in this.orderedPropertyMaps)
            {
                var cellReference = this.ConstructExcelCellReference(this.currentRowIndex, propertyMap.PropertyData.IndexWrite);
                var columnType = this.columnTypes[propertyMap.PropertyData.IndexWrite];

                var cellValue = this.ResolveCellValue(propertyMap, record);
                var cellAttributes = new List<OpenXmlAttribute>() { new OpenXmlAttribute(styleIndexAttribute, null, this.columnStyles[propertyMap.PropertyData.IndexWrite].ToString()) };
                this.WriteCellValue(cellValue, cellReference, columnType, cellAttributes);

                if (this.worksheetStyle.ShouldAutoFitColumns)
                {
                    this.AddColumnWidth(propertyMap.PropertyData.IndexWrite, cellValue.Length);
                }
            }

            this.writer.WriteEndElement(); // row
            this.currentRowIndex++;
        }

        /// <summary>
        /// Writes the provided records at the current position.
        /// </summary>
        /// <param name="records">The records to be written.</param>
        public void WriteRecords(IEnumerable<TClass> records)
        {
            foreach (var record in records)
            {
                this.WriteRecord(record);
            }
        }

        /// <summary>
        /// Converts between a numeric row and column index to an Excel cell position (e.g., A1).
        /// </summary>
        /// <param name="rowIndex">The row index.</param>
        /// <param name="columnIndex">The column index.</param>
        /// <returns>An Excel cell reference.</returns>
        private string ConstructExcelCellReference(uint rowIndex, uint columnIndex)
        {
            string columnName = string.Empty;
            if (this.columnCellReferences.TryGetValue(columnIndex, out columnName))
            {
                return columnName + rowIndex.ToString();
            }

            uint dividend = columnIndex;
            uint modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            this.columnCellReferences.Add(columnIndex, columnName);
            return columnName + rowIndex.ToString();
        }

        private static List<OpenXmlAttribute> ConstructRowAttributes(uint rowIndex) => new List<OpenXmlAttribute> { new OpenXmlAttribute(rowIndexAttribute, null, rowIndex.ToString()) };

        private static string ConvertBoolToOpenXmlFormat(bool value) => value ? "1" : "0";

        private static string ConvertDateTimeToOpenXmlFormat(in DateTime dateTime) => dateTime.ToOADate().ToString(CultureInfo.InvariantCulture);

        private Sheet ConstructSheetPart()
        {
            var sheets = this.spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            return new Sheet()
            {
                Id = this.spreadsheetDocument.WorkbookPart.GetIdOfPart(this.worksheetPart),
                Name = this.worksheetName,
                SheetId = sheetId,
            };
        }

        private SheetViews ConstructSheetViews()
        {
            var sheetViews = new SheetViews();
            var sheetView = new SheetView()
            {
                WorkbookViewId = 0U
            };

            if (this.worksheetStyle.ShouldFreezeHeaderRow)
            {
                var firstFrozenCellReference = this.ConstructExcelCellReference(this.worksheetStyle.HeaderRowIndex + 1, 1);
                var pane = new Pane()
                {
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen,
                    TopLeftCell = firstFrozenCellReference,
                    VerticalSplit = this.worksheetStyle.HeaderRowIndex,
                };

                var selection = new Selection()
                {
                    ActiveCell = firstFrozenCellReference,
                    Pane = PaneValues.BottomLeft,
                };

                sheetView.Append(pane);
                sheetView.Append(selection);
            }

            sheetViews.Append(sheetView);

            return sheetViews;
        }

        private IEnumerable<PropertyMap<TClass>> CreateOrderedPropertyMaps()
        {
            var classMap = Activator.CreateInstance<TClassMap>();
            var propertyMaps = classMap.PropertyMaps.Where(x => !x.PropertyData.IgnoreWrite);
            var indexes = new HashSet<uint>();

            uint currentIndex = 1;
            foreach (var map in propertyMaps.Where(x => x.PropertyData.IndexWrite > 0))
            {
                indexes.Add(map.PropertyData.IndexWrite);
            }

            foreach (var map in propertyMaps.Where(x => x.PropertyData.IndexWrite == 0))
            {
                while (true)
                {
                    if (!indexes.Contains(currentIndex))
                    {
                        map.PropertyData.IndexWrite = currentIndex;
                        indexes.Add(currentIndex);
                        currentIndex++;
                        break;
                    }

                    currentIndex++;
                }
            }

            return propertyMaps.OrderBy(x => x.PropertyData.IndexWrite);
        }

        private string InsertSharedString(string text)
        {
            bool stringExists = this.sharedStrings.TryGetValue(text, out string sharedStringIndex);
            if (stringExists)
            {
                return sharedStringIndex;
            }

            var newIndex = this.sharedStrings.Count.ToString();
            this.sharedStrings.Add(text, newIndex);

            return newIndex;
        }

        private static string ResolveHeader(PropertyMap<TClass> propertyMap)
        {
            if (string.IsNullOrWhiteSpace(propertyMap.PropertyData.NameWrite))
            {
                // account for columns not mapped to a property (e.g., constants) with no name specified
                return propertyMap.PropertyData.Property == null ? string.Empty : propertyMap.PropertyData.Property.Name;
            }

            return propertyMap.PropertyData.NameWrite;
        }

        private string ResolveCellValue(PropertyMap<TClass> propertyMap, TClass record)
        {
            var columnType = this.columnTypes[propertyMap.PropertyData.IndexWrite];

            object value;

            if (propertyMap.PropertyData.WriteUsing != null)
            {
                value = propertyMap.PropertyData.WriteUsing(record);
            }
            else if (propertyMap.PropertyData.ConstantWrite != null)
            {
                value = propertyMap.PropertyData.ConstantWrite;
            }
            else
            {
                value = propertyMap.PropertyData.Property.GetValue(record);
                if ( value == null || (propertyMap.PropertyData.Property.PropertyType == typeof(char) && (char)value == char.MinValue))
                {
                    value = propertyMap.PropertyData.DefaultWrite ?? string.Empty;
                }
            }

            switch (columnType)
            {
                case CellValues.Boolean:
                    return ConvertBoolToOpenXmlFormat((bool)value);

                case CellValues.Date:
                    return ConvertDateTimeToOpenXmlFormat((DateTime)value);

                case CellValues.Number:
                    return string.IsNullOrWhiteSpace(value.ToString()) ? string.Empty : Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString();

                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        private void WriteCellValue(string cellValue, string cellReference, CellValues cellType, List<OpenXmlAttribute> attributes)
        {
            attributes.Add(cellValueAttributes[cellType]);

            cellReferenceAttribute.Value = cellReference;
            attributes.Add(cellReferenceAttribute);

            this.writer.WriteStartElement(new Cell(), attributes);

            switch (cellType)
            {
                case CellValues.InlineString:
                    var parsedInlineString = new InlineString(new Text(cellValue));
                    this.writer.WriteElement(parsedInlineString);
                    break;

                case CellValues.SharedString:
                    var sharedStringIndex = this.InsertSharedString(cellValue);
                    this.writer.WriteElement(new CellValue(sharedStringIndex));
                    break;

                default:
                    this.writer.WriteElement(new CellValue(cellValue));
                    break;
            }

            this.writer.WriteEndElement();
        }

        private void WriteFinalWorksheetElements()
        {
            this.writer.WriteEndElement(); // SheetData

            if (this.worksheetStyle.ShouldAutoFilter)
            {
				var firstCellReference = this.ConstructExcelCellReference(this.worksheetStyle.HeaderRowIndex, 1);
				var lastIndex = this.orderedPropertyMaps.Last().PropertyData.IndexWrite;
				var lastCellReference = this.ConstructExcelCellReference(this.worksheetStyle.HeaderRowIndex, lastIndex);
                this.writer.WriteElement(new AutoFilter() { Reference = $"{firstCellReference}:{lastCellReference}" });
            }

            this.writer.WriteEndElement(); // Worksheet
        }

        private void WriteInitialWorksheetElements()
        {
            this.writer.WriteStartElement(new Worksheet());
            this.writer.WriteElement(this.ConstructSheetViews());
            this.WriteColumnsPart(this.writer);
            this.writer.WriteStartElement(new SheetData());

            if (this.worksheetStyle.ShouldWriteHeaderRow)
            {
                this.WriteHeader();
            }
        }

        #region columns

        private void AddColumnWidth(uint columnIndex, int textLength)
        {
            if (this.columnWidths.ContainsKey(columnIndex))
            {
                if (textLength > this.columnWidths[columnIndex])
                {
                    this.columnWidths[columnIndex] = textLength;
                }
            }
            else
            {
                this.columnWidths.Add(columnIndex, textLength);
            }
        }

        private static double CalculateColumnWidth(int textLength, System.Drawing.Font font)
        {
            const int cellPadding = 0;
            using (var graphics = Graphics.FromImage(new Bitmap(200, 200)))
            {
                float zeroDigitWidth = graphics.MeasureString('0'.ToString(), font).Width;
                float desiredWidth = (textLength * zeroDigitWidth) + cellPadding;
                return Math.Truncate(desiredWidth / zeroDigitWidth * 256) / 256;
            }
        }

        private void ImplementAutoFitColumns()
        {
            var replacementWorksheetPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            using (var reader = OpenXmlReader.Create(this.worksheetPart))
            using (var replacementWriter = OpenXmlWriter.Create(replacementWorksheetPart))
            {
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Columns))
                    {
                        if (reader.IsEndElement)
                        {
                            this.WriteColumnsPart(replacementWriter);
                        }
                    }
                    else if (reader.ElementType == typeof(Column))
                    {
                        continue;
                    }
                    else
                    {
                        if (reader.ElementType == typeof(CellValue))
                        {
                            if (reader.IsStartElement)
                            {
                                replacementWriter.WriteStartElement(reader);
                                replacementWriter.WriteString(reader.GetText());
                            }
                            else if (reader.IsEndElement)
                            {
                                replacementWriter.WriteEndElement();
                            }
                        }
                        else if (reader.IsStartElement)
                        {
                            replacementWriter.WriteStartElement(reader);
                        }
                        else if (reader.IsEndElement)
                        {
                            replacementWriter.WriteEndElement();
                        }
                    }
                }
            }

            this.spreadsheetDocument.WorkbookPart.DeletePart(this.worksheetPart);
            this.worksheetPart = replacementWorksheetPart;
        }

        private static double ResolveColumnWidth(double actualWidth, double minWidth, double maxWidth) =>
            actualWidth > maxWidth ? maxWidth : actualWidth < minWidth ? minWidth : actualWidth;

        private void WriteColumnsPart(OpenXmlWriter openXmlWriter)
        {
            const int filterButtonWidth = 3;

            openXmlWriter.WriteStartElement(new Columns());
            foreach (var propertyMap in this.orderedPropertyMaps)
            {
                if (!this.columnWidths.TryGetValue(propertyMap.PropertyData.IndexWrite, out var textLength))
                {
                    textLength = 0;
                }

                if (this.worksheetStyle.ShouldAutoFilter)
                {
                    textLength += filterButtonWidth;
                }

                var calculatedColumnWidth = CalculateColumnWidth(textLength, propertyMap.PropertyData.Style.Font);
                var resolvedColumnWidth = ResolveColumnWidth(calculatedColumnWidth, this.worksheetStyle.MinColumnWidth, this.worksheetStyle.MaxColumnWidth);
                var columnIndexText = propertyMap.PropertyData.IndexWrite.ToString();

                openXmlWriter.WriteStartElement(new Column(), new List<OpenXmlAttribute>()
                {
                    new OpenXmlAttribute("min", null, columnIndexText),
                    new OpenXmlAttribute("max", null, columnIndexText),
                    new OpenXmlAttribute("width", null, resolvedColumnWidth.ToString()),
                    new OpenXmlAttribute("bestFit", null, ConvertBoolToOpenXmlFormat(true)),
                    new OpenXmlAttribute("customWidth", null, ConvertBoolToOpenXmlFormat(true)),
                });

                openXmlWriter.WriteEndElement(); // column
            }

            openXmlWriter.WriteEndElement(); // columns
        }

        #endregion columns

        #region cache data

        private void CacheColumnStyles()
        {
            const uint defaultCellStyleFormatId = 0;

            foreach (var propertyMap in this.orderedPropertyMaps)
            {
                uint borderId = this.spreadsheetStyles.AddBorder(propertyMap.PropertyData.Style.BorderPlacement, propertyMap.PropertyData.Style.BorderStyle, propertyMap.PropertyData.Style.BorderColor);
                uint fillId = this.spreadsheetStyles.AddPatternFill(propertyMap.PropertyData.Style.BackgroundColor, propertyMap.PropertyData.Style.BackgroundPatternType);
                uint fontId = this.spreadsheetStyles.AddFont(propertyMap.PropertyData.Style.Font, propertyMap.PropertyData.Style.ForegroundColor);

                uint numberFormatId;
                if (propertyMap.PropertyData.Style.IsNumberFormatSpecified)
                {
                    numberFormatId = string.IsNullOrWhiteSpace(propertyMap.PropertyData.Style.CustomNumberFormat) ?
                        (uint)propertyMap.PropertyData.Style.NumberFormat :
                        this.spreadsheetStyles.AddNumberingFormat(propertyMap.PropertyData.Style.CustomNumberFormat);
                }
                else
                {
                    var columnType = this.columnTypes[propertyMap.PropertyData.IndexWrite];
                    numberFormatId = columnType == CellValues.Date ? (uint)OpenXmlNumberingFormat.DateMonthDayYear : (uint)OpenXmlNumberingFormat.General;
                }

                uint styleId = this.spreadsheetStyles.AddCellFormat(borderId, fillId, fontId, defaultCellStyleFormatId, numberFormatId, propertyMap.PropertyData.Style.HoizontalAlignment, propertyMap.PropertyData.Style.VerticalAlignment);
                this.columnStyles.Add(propertyMap.PropertyData.IndexWrite, styleId);
            }
        }

        private void CacheColumnTypes()
        {
            foreach (var propertyMap in this.orderedPropertyMaps)
            {
                switch (propertyMap.PropertyData.ColumnType)
                {
                    case ColumnType.Boolean:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Boolean);
                        break;

                    case ColumnType.Number:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Number);
                        break;

                    case ColumnType.Date:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Date);
                        break;

                    case ColumnType.Formula:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.String);
                        break;

                    case ColumnType.RichText:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.InlineString);
                        break;

                    case ColumnType.Text:
                        this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.SharedString);
                        break;

                    case ColumnType.Unset:

                        var propertyType =
                            propertyMap.PropertyData.Property == null ?
                            propertyMap.PropertyData.ConstantWrite.GetType() :
                            propertyMap.PropertyData.Property.PropertyType;

                        if (propertyType == typeof(bool))
                        {
                            this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Boolean);
                            propertyMap.ColumnType(ColumnType.Boolean);
                        }
                        else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTimeOffset))
                        {
                            this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Date);
                        }
                        else if (propertyType.IsNumeric() || Nullable.GetUnderlyingType(propertyType).IsNumeric())
                        {
                            this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.Number);
                        }
                        else
                        {
                            this.columnTypes.Add(propertyMap.PropertyData.IndexWrite, CellValues.SharedString);
                        }

                        break;
                }
            }
        }

        private void CacheHeaderStyles()
        {
            uint headerFillId = this.spreadsheetStyles.AddPatternFill(this.worksheetStyle.HeaderBackgroundColor, this.worksheetStyle.HeaderBackgroundPatternType);
            uint headerFontId = this.spreadsheetStyles.AddFont(this.worksheetStyle.HeaderFont, this.worksheetStyle.HeaderForegroundColor);
            uint headerStyleId = this.spreadsheetStyles.AddCellFormat(fillId: headerFillId, fontId: headerFontId, horizontalAlignment: this.worksheetStyle.HeaderHoizontalAlignment, verticalAlignment: this.worksheetStyle.HeaderVerticalAlignment);

            this.columnStyles.Add(headerStyleKey, headerStyleId);
        }

        private void CacheWorksheetData()
        {
            this.CacheHeaderStyles();
            this.CacheColumnTypes();
            this.CacheColumnStyles();
        }

        #endregion cache data

        #region IDisposable

        /// <summary>
        /// Closes the <see cref="WorksheetWriter{TClass, TClassMap}"/> object and the underlying stream, writes any 
        /// remaining worksheet data, and releases any the system resources associated with the writer.
        /// </summary>
        public void Close()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="WorksheetWriter{TClass, TClassMap}"/> object and the underlying stream, writes any 
        /// remaining worksheet data, and releases any the system resources associated with the writer.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="WorksheetWriter{TClass, TClassMap}"/> object and the underlying stream, writes any 
        /// remaining worksheet data, and optionally releases any the system resources associated with the writer.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.isDisposed && disposing)
            {
                this.WriteFinalWorksheetElements();
                this.writer.Close();

                if (this.worksheetStyle.ShouldAutoFitColumns)
                {
                    this.ImplementAutoFitColumns();
                }

                var sheet = this.ConstructSheetPart();
                this.spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().InsertAt(sheet, this.worksheetPositionIndex);
            }

            this.isDisposed = true;
        }

        #endregion IDisposable
    }
}