namespace OpenSpreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using OpenSpreadsheet.Configuration;

    /// <summary>
    /// Writes data to a worksheet.
    /// </summary>
    /// <typeparam name="TClass">The class type to be written.</typeparam>
    /// <typeparam name="TClassMap">A map defining how to write record data to the worksheet.</typeparam>
    public class WorksheetReader<TClass, TClassMap> : IDisposable
        where TClass : class, new()
        where TClassMap : ClassMap<TClass>
    {
        private const string rowIndexAttribute = "r";

        private readonly Dictionary<string, uint> columnCellReferences = new Dictionary<string, uint>();
        private readonly IReadOnlyCollection<PropertyMap<TClass>> propertyMaps;
        private readonly OpenXmlReader reader;
        private readonly BidirectionalDictionary<string, string> sharedStrings;
        private readonly SpreadsheetDocument spreadsheetDocument;

        private uint currentRowIndex = 1;
        private bool isDisposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetReader{TClass, TClassMap}"/> class.
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sharedStrings"></param>
        /// <param name="headerRowIndex"></param>
        public WorksheetReader(string worksheetName, SpreadsheetDocument spreadsheetDocument, BidirectionalDictionary<string, string> sharedStrings, uint headerRowIndex = 1)
        {
            this.sharedStrings = sharedStrings;
            this.spreadsheetDocument = spreadsheetDocument;

            // reader setup
            var worksheetId = this.spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().First(s => worksheetName.Equals(s.Name)).Id;
            var worksheetPart = this.spreadsheetDocument.WorkbookPart.GetPartById(worksheetId);
            this.reader = OpenXmlReader.Create(worksheetPart);
            this.reader.Read();
            this.ReadHeader(headerRowIndex);

            // map setup
            var classMap = Activator.CreateInstance<TClassMap>();
            this.propertyMaps = classMap.PropertyMaps.Where(x => !x.PropertyData.IgnoreRead).ToList().AsReadOnly();
        }

        /// <summary>
        /// Gets a collection of key value pairs containing a header column index and its name.
        /// </summary>
        public BidirectionalDictionary<uint, string> Headers { get; } = new BidirectionalDictionary<uint, string>();

        /// <summary>
        /// Read a single row at the current position and map its data to an object.
        /// </summary>
        /// <returns>A mapped object.</returns>
        public TClass ReadRow()
        {
            var readerRow = this.ReadRowValues();
            if (readerRow == null)
            {
                return null;
            }

            var record = new TClass();
            foreach (var map in this.propertyMaps)
            {
                var propertyInfo = record.GetType().GetProperty(map.PropertyData.Property.Name);
                if (!propertyInfo.CanWrite)
                {
                    continue;
                }

                if (map.PropertyData.ConstantRead != null)
                {
                    propertyInfo.SetValue(record, map.PropertyData.ConstantRead);
                    continue;
                }

                if (map.PropertyData.ReadUsing != null)
                {
                    var expressionValue = map.PropertyData.ReadUsing(readerRow);
                    propertyInfo.SetValue(record, expressionValue);
                    continue;
                }

                uint columnIndex;
                if (map.PropertyData.IndexRead > 0)
                {
                    columnIndex = map.PropertyData.IndexRead;
                }
                else
                {
                    var cellName = string.IsNullOrWhiteSpace(map.PropertyData.NameRead)
                        ? map.PropertyData.Property.Name
                        : map.PropertyData.NameRead;

                    if (!this.Headers.TryGetKey(cellName, out columnIndex))
                    {
                        throw new InvalidOperationException($"The ClassMap contains invalid read maps. The property {map.PropertyData.Property.Name} has no index defined and there is no spreadsheet column matching either the property name or a defined name property.");
                    }
                }

                var cellValue = readerRow.GetCellValue(columnIndex);

                if (map.PropertyData.DefaultRead != null && cellValue.Length == 0)
                {
                    propertyInfo.SetValue(record, map.PropertyData.DefaultRead);
                    continue;
                }

                var propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;

                object safeValue;
                if (propertyType == typeof(bool))
                {
                    safeValue = (cellValue == null) ? null : (object)ConvertStringToBool(cellValue);
                }
                else if (propertyType == typeof(DateTime))
                {
                    safeValue = (cellValue == null) ? null : (object)ConvertDateTime(cellValue);
                }
                else if (propertyType.IsEnum)
                {
                    safeValue = Enum.Parse(propertyType, cellValue);
                }
                else
                {
                    safeValue = (cellValue == null) ? null : Convert.ChangeType(cellValue, propertyType);
                }

                propertyInfo.SetValue(record, safeValue, null);
            }

            return record;
        }

        /// <summary>
        /// Read all rows starting from the current position and map the data to a collection of objects.
        /// </summary>
        /// <returns>A collection of mapped objects.</returns>
        public IEnumerable<TClass> ReadRows()
        {
            var records = new List<TClass>();
            do
            {
                var record = this.ReadRow();
                if (record != null)
                {
                    records.Add(record);
                }
            } while (!this.reader.EOF);

            return records;
        }

        /// <summary>
        /// Skip a single row.
        /// </summary>
        public void SkipRow() => this.SkipRows(1);

        /// <summary>
        /// Skip one or more rows.
        /// </summary>
        /// <param name="count">The number of rows to skip.</param>
        public void SkipRows(uint count)
        {
            var targetRow = this.currentRowIndex + count;
            do
            {
                this.AdvanceToRowStart();
                if (this.currentRowIndex < targetRow)
                {
                    this.reader.Skip();
                }
                else
                {
                    return;
                }
            } while (!this.reader.EOF);
        }

        private void AdvanceToRowStart()
        {
            while (!this.reader.EOF)
            {
                if (this.reader.ElementType == typeof(Row) && this.reader.IsStartElement)
                {
                    this.currentRowIndex = uint.Parse(this.reader.Attributes.First(r => r.LocalName == rowIndexAttribute).Value);
                    return;
                }
                else
                {
                    this.reader.Read();
                }
            }
        }

        private static DateTime ConvertDateTime(string date)
        {
            if (DateTime.TryParse(date, out DateTime datetimeResult))
            {
                return datetimeResult;
            }

            if (double.TryParse(date, out double doubleResult))
            {
                return DateTime.FromOADate(doubleResult);
            }

            throw new InvalidCastException();
        }

        private static bool ConvertStringToBool(string textBool)
        {
            if (bool.TryParse(textBool, out bool boolValue))
            {
                return boolValue;
            }

            if (int.TryParse(textBool, out int intValue))
            {
                return (bool)Convert.ChangeType(intValue, typeof(bool));
            }

            throw new InvalidCastException();
        }

        private static string GetCellValue(BidirectionalDictionary<string, string> sharedStrings, Cell cell)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            if (cell.DataType == CellValues.SharedString)
            {
                sharedStrings.TryGetKey(cell.CellValue.InnerText, out string sharedStringValue);
                return sharedStringValue;
            }

            return cell.CellValue.InnerText;
        }

        /// <summary>
        /// Determines a cell's one-based column index from its Excel cell position (e.g., A1).
        /// </summary>
        /// <param name="cellReference">The cell reference to be converted.</param>
        /// <returns>The cell's numeric column index.</returns>
        private uint GetColumnIndexFromCellReference(string cellReference)
        {
            var columnLetters = cellReference.Where(c => !char.IsNumber(c)).ToArray();
            string columnReference = new string(columnLetters);
            if (this.columnCellReferences.ContainsKey(columnReference))
            {
                return this.columnCellReferences[columnReference];
            }

            int sum = 0;
            for (int i = 0; i < columnLetters.Length; i++)
            {
                sum *= 26;
                sum += columnLetters[i] - 'A' + 1;
            }

            this.columnCellReferences.Add(columnReference, (uint)sum);
            return (uint)sum;
        }

        private void ReadHeader(uint headerRowIndex)
        {
            if (headerRowIndex == 0)
            {
                return;
            }

            this.SkipRows(headerRowIndex - 1);

            this.AdvanceToRowStart();
            if (this.reader.EOF)
            {
                throw new InvalidOperationException("There are no rows available to read.");
            }

            this.reader.ReadFirstChild();
            do
            {
                if (this.reader.ElementType == typeof(Cell))
                {
                    var cell = (Cell)this.reader.LoadCurrentElement();
                    var cellValue = GetCellValue(this.sharedStrings, cell);
                    var columnIndex = this.GetColumnIndexFromCellReference(cell.CellReference);

                    this.Headers.Add(columnIndex, cellValue);
                }
            } while (this.reader.ReadNextSibling());
        }

        private ReaderRow ReadRowValues()
        {
            this.AdvanceToRowStart();
            if (this.reader.EOF)
            {
                return null;
            }

            this.reader.ReadFirstChild();
            var rowValues = new Dictionary<uint, string>();

            do
            {
                if (this.reader.ElementType == typeof(Cell))
                {
                    var cell = (Cell)this.reader.LoadCurrentElement();
                    var cellValue = GetCellValue(this.sharedStrings, cell);
                    var columnIndex = this.GetColumnIndexFromCellReference(cell.CellReference);

                    rowValues.Add(columnIndex, cellValue);
                }
            } while (this.reader.ReadNextSibling());

            return new ReaderRow(this.Headers, rowValues);
        }

        #region IDisposable

        /// <summary>
        /// Closes the <see cref="WorksheetReader{TClass, TClassMap}"/> object and the underlying stream, and releases any the system resources associated with the reader.
        /// </summary>
        public void Close()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="WorksheetReader{TClass, TClassMap}"/> object and the underlying stream, and releases any the system resources associated with the reader.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="WorksheetReader{TClass, TClassMap}"/> object and the underlying stream, and optionally releases any the system resources associated with the reader.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.isDisposed && disposing)
            {
                this.reader.Close();
            }

            this.isDisposed = true;
        }

        #endregion IDisposable
    }
}