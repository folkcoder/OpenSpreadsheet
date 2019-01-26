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
        private readonly Dictionary<uint, PropertyMap> propertyMaps;
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
            this.Headers = ReadHeader(headerRowIndex);

            // map setup
            this.propertyMaps = CreateOrderedPropertyMaps();
        }

        /// <summary>
        /// Gets a collection of key value pairs containing a header column index and its name.
        /// </summary>
        public Dictionary<uint, string> Headers { get; }

        private Dictionary<uint, string> ReadHeader(uint headerRowIndex)
        {
            if (headerRowIndex == 0)
            {
                return null;
            }

            this.SkipRows(headerRowIndex - 1);

            var headers = new Dictionary<uint, string>();
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
                    var columnIndex = GetColumnIndexFromCellReference(cell.CellReference);

                    headers.Add(columnIndex, cellValue);
                }
            } while (this.reader.ReadNextSibling());

            return headers;
        }

        /// <summary>
        /// Read a single row at the current position and map its data to an object.
        /// </summary>
        /// <returns>A mapped object.</returns>
        public TClass ReadRow()
        {
            this.AdvanceToRowStart();
            if (this.reader.EOF)
            {
                return null;
            }

            var record = new TClass();
            this.reader.ReadFirstChild();

            do
            {
                if (this.reader.ElementType == typeof(Cell))
                {
                    var cell = (Cell)this.reader.LoadCurrentElement();
                    var cellValue = GetCellValue(this.sharedStrings, cell);
                    var columnIndex = GetColumnIndexFromCellReference(cell.CellReference);
                    if (this.propertyMaps.ContainsKey(columnIndex))
                    {
                        var propertyMap = this.propertyMaps[columnIndex];
                        var propertyInfo = record.GetType().GetProperty(propertyMap.PropertyData.Property.Name);

                        if (propertyMap.PropertyData.ConstantRead != null)
                        {
                            propertyInfo.SetValue(record, propertyMap.PropertyData.ConstantRead);
                        }
                        else
                        {
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
                            else
                            {
                                safeValue = (cellValue == null) ? null : Convert.ChangeType(cellValue, propertyType);
                            }

                            if (safeValue == null && propertyMap.PropertyData.DefaultRead != null)
                            {
                                safeValue = propertyMap.PropertyData.DefaultRead;
                            }

                            propertyInfo.SetValue(record, safeValue, null);
                        }
                    }
                }
            } while (this.reader.ReadNextSibling());

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
        public void SkipRow()
        {
            this.SkipRows(1);
        }

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
            else
            {
                throw new InvalidCastException();
            }
        }

        private Dictionary<uint, PropertyMap> CreateOrderedPropertyMaps()
        {
            var indexes = new Dictionary<uint, PropertyMap>();

            var classMap = Activator.CreateInstance<TClassMap>();
            var propertyMaps = classMap.PropertyMaps.Where(x => !x.PropertyData.IgnoreRead);

            foreach (var map in propertyMaps.Where(x => x.PropertyData.IndexRead > 0))
            {
                indexes.Add(map.PropertyData.IndexRead, map);
            }

            var mapsWithUndefinedIndexes = propertyMaps.Where(x => x.PropertyData.IndexRead == 0 && x.PropertyData.ConstantRead == null);
            if (mapsWithUndefinedIndexes != null && this.Headers == null)
            {
                throw new InvalidOperationException("The ClassMap contains invalid read maps. Each property map must define either a column index or a header name.");
            }

            foreach (var map in mapsWithUndefinedIndexes)
            {
                string mapName = map.PropertyData.NameRead ?? map.PropertyData.Property.Name;
                var matchedColumnIndex = this.Headers.FirstOrDefault(x => x.Value.Equals(mapName, StringComparison.InvariantCultureIgnoreCase));

                if (matchedColumnIndex.Value == null)
                {
                    throw new InvalidOperationException($"The ClassMap contains invalid read maps. The property {map.PropertyData.Property.Name} has no index defined and there is no spreadsheet column matching either the property name or a defined name property.");
                }

                map.PropertyData.IndexRead = matchedColumnIndex.Key;
                indexes.Add(map.PropertyData.IndexRead, map);
            }

            return indexes;
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
                sum += (columnLetters[i] - 'A' + 1);
            }

            this.columnCellReferences.Add(columnReference, (uint)sum);
            return (uint)sum;
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