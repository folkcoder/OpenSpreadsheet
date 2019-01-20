namespace SpreadsheetHelper
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using SpreadsheetHelper.Configuration;

    /// <summary>
    /// Writes data to a worksheet.
    /// </summary>
    /// <typeparam name="TClass">The class type to be written.</typeparam>
    /// <typeparam name="TClassMap">A map defining how to write record data to the worksheet.</typeparam>
    public class WorksheetReader<TClass, TClassMap> : IDisposable
        where TClass : class
        where TClassMap : ClassMap<TClass>
    {
        private const string rowIndexAttribute = "r";
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

            // create default instance of class, supplying defined optional arguments if applicable
            var classMap = Activator.CreateInstance<TClassMap>();
            this.propertyMaps = classMap.PropertyMaps
                .Where(x => !x.PropertyData.IgnoreRead)
                .ToDictionary(x => x.PropertyData.IndexRead, x => x);

            // reader setup
            var worksheetId = this.spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().First(s => worksheetName.Equals(s.Name)).Id;
            var worksheetPart = this.spreadsheetDocument.WorkbookPart.GetPartById(worksheetId);
            this.reader = OpenXmlReader.Create(worksheetPart);
            this.reader.Read();
            this.SkipRows(headerRowIndex);
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

            // TODO: Allow for optional constructor parameters
            var record = Activator.CreateInstance<TClass>();

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
        private static uint GetColumnIndexFromCellReference(string cellReference)
        {
            var columnLetters = cellReference.Where(c => !char.IsNumber(c)).ToArray();
            int sum = 0;

            for (int i = 0; i < columnLetters.Length; i++)
            {
                sum *= 26;
                sum += (columnLetters[i] - 'A' + 1);
            }

            return (uint)sum;
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