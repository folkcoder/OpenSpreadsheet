namespace SpreadsheetHelper
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using SpreadsheetHelper.Configuration;

    /// <summary>
    /// Provides a wrapper around the Open XML SDF to more easily create spreadsheet files with one or more worksheets.
    /// </summary>
    public class Spreadsheet : IDisposable
    {
        private bool isDisposed = false;
        private int originalSharedStringCount = 0;

        private readonly BidirectionalDictionary<string, string> sharedStrings = new BidirectionalDictionary<string, string>();
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly StylesCollection stylesheet = new StylesCollection();
        private readonly HashSet<Type> validatedClassMaps = new HashSet<Type>();
        private readonly HashSet<string> worksheetNames = new HashSet<string>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// </summary>
        /// <param name="filePath">The full path and file name of the XLSX file to be generated.</param>
        public Spreadsheet(string filePath)
        {
            if (File.Exists(filePath))
            {
                this.spreadsheetDocument = SpreadsheetDocument.Open(filePath, true);
                this.LoadSharedStringTable();
                this.LoadWorkbookStyles();
            }
            else
            {
                this.spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook, true);
                this.spreadsheetDocument.AddWorkbookPart();
                WriteWorkbookPart(this.spreadsheetDocument.WorkbookPart);
            }
        }

        /// <summary>
        /// Creates a new worksheet reader.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be read from the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be read.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be read.</param>
        /// <param name="headerRowIndex">The one-based row index of the header row. A value of 0 indicates a worksheet with no header.</param>
        /// <returns>A new worksheet reader.</returns>
        public WorksheetReader<TClass, TClassMap> CreateWorksheetReader<TClass, TClassMap>(string worksheetName, uint headerRowIndex = 1)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            this.ValidateClassMap<TClass, TClassMap>();
            return new WorksheetReader<TClass, TClassMap>(worksheetName, this.spreadsheetDocument, this.sharedStrings, headerRowIndex);
        }

        /// <summary>
        /// Creates a new worksheet writer.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be written to the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be written.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be written.</param>
        /// <param name="worksheetStyle">The worksheet's style definition.</param>
        /// <returns>A new worksheet writer.</returns>
        public WorksheetWriter<TClass, TClassMap> CreateWorksheetWriter<TClass, TClassMap>(string worksheetName, WorksheetStyle worksheetStyle = null)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            this.ValidateClassMap<TClass, TClassMap>();
            this.AddWorksheetName(worksheetName);
            if (worksheetStyle == null)
            {
                worksheetStyle = new WorksheetStyle();
            }

            var positionIndex = this.spreadsheetDocument.WorkbookPart.Workbook.Sheets.Count();
            return new WorksheetWriter<TClass, TClassMap>(worksheetName, positionIndex, this.spreadsheetDocument, worksheetStyle, this.sharedStrings, this.stylesheet);
        }

        /// <summary>
        /// Creates a new worksheet writer.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be written to the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be written.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be written.</param>
        /// <param name="worksheetIndex">The worksheet's zero-based position index within the workbook.</param>
        /// <param name="worksheetStyle">The worksheet's style definition.</param>
        /// <returns>A new worksheet writer.</returns>
        public WorksheetWriter<TClass, TClassMap> CreateWorksheetWriter<TClass, TClassMap>(string worksheetName, int worksheetIndex, WorksheetStyle worksheetStyle = null)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            this.ValidateClassMap<TClass, TClassMap>();
            this.AddWorksheetName(worksheetName);
            if (worksheetStyle == null)
            {
                worksheetStyle = new WorksheetStyle();
            }

            return new WorksheetWriter<TClass, TClassMap>(worksheetName, worksheetIndex, this.spreadsheetDocument, worksheetStyle, this.sharedStrings, this.stylesheet);
        }

        /// <summary>
        /// Reads the worksheet with the provided worksheet name and maps its data to a typed collection.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be read from the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be read.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be read.</param>
        /// <param name="headerRowIndex">The one-based row index of the header row. A value of 0 indicates a worksheet with no header.</param>
        /// <returns>A collection containing the worksheet's mapped data.</returns>
        public IEnumerable<TClass> ReadWorksheet<TClass, TClassMap>(string worksheetName, uint headerRowIndex = 1)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            using (var worksheetReader = this.CreateWorksheetReader<TClass, TClassMap>(worksheetName, headerRowIndex))
            {
                return worksheetReader.ReadRows();
            }
        }

        /// <summary>
        /// Writes the provided data to a new worksheet.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be written to the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be written.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be written.</param>
        /// <param name="records">The data to be written.</param>
        /// <param name="worksheetStyle">The worksheet's style definition.</param>
        public void WriteWorksheet<TClass, TClassMap>(string worksheetName, IEnumerable<TClass> records, WorksheetStyle worksheetStyle = null)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            using (var worksheetWriter = this.CreateWorksheetWriter<TClass, TClassMap>(worksheetName, worksheetStyle))
            {
                worksheetWriter.WriteRecords(records);
            }
        }

        /// <summary>
        /// Writes the provided data to a new worksheet.
        /// </summary>
        /// <typeparam name="TClass">The class type encapsulating the data to be written to the worksheet.</typeparam>
        /// <typeparam name="TClassMap">A class map identifying how the data will be written.</typeparam>
        /// <param name="worksheetName">The name of the worksheet to be written.</param>
        /// <param name="worksheetIndex">The worksheet's zero-based position index within the workbook.</param>
        /// <param name="records">The data to be written.</param>
        /// <param name="worksheetStyle">The worksheet's style definition.</param>
        public void WriteWorksheet<TClass, TClassMap>(string worksheetName, int worksheetIndex, IEnumerable<TClass> records, WorksheetStyle worksheetStyle = null)
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            using (var worksheetWriter = this.CreateWorksheetWriter<TClass, TClassMap>(worksheetName, worksheetIndex, worksheetStyle))
            {
                worksheetWriter.WriteRecords(records);
            }
        }

        private static void WriteWorkbookPart(WorkbookPart workbookPart)
        {
            using (var writer = OpenXmlWriter.Create(workbookPart))
            {
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }

        private void AddWorksheetName(string worksheetName)
        {
            if (this.worksheetNames.Contains(worksheetName))
            {
                throw new ArgumentException($"A worksheet with the name '{worksheetName}' has already been added to the spreadsheet. Worksheet names must be unique.");
            }

            var invalidChars = new char[] { '\\', '/', '*', '?', ':', '[', ']' };
            if (worksheetName.Length > 31 || worksheetName.Any(c => invalidChars.Contains(c)))
            {
                throw new ArgumentException($"The worksheet named '{worksheetName}' contains one of the following invalid characters: {invalidChars}.");
            }

            this.worksheetNames.Add(worksheetName);
        }

        private void LoadSharedStringTable()
        {
            var sharedStringTablePart = this.spreadsheetDocument.WorkbookPart.SharedStringTablePart;
            if (sharedStringTablePart == null)
            {
                return;
            }

            using (var reader = OpenXmlReader.Create(sharedStringTablePart.SharedStringTable))
            {
                int i = 0;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(SharedStringItem))
                    {
                        var sharedStringItem = (SharedStringItem)reader.LoadCurrentElement();
                        this.sharedStrings.Add(sharedStringItem.Text != null ? sharedStringItem.Text.Text : string.Empty, i.ToString());
                        i++;
                    }
                }

                this.originalSharedStringCount = i;
            }
        }

        private void LoadWorkbookStyles()
        {
            var workbookStylesPart = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart;
            if (workbookStylesPart == null)
            {
                return;
            }

            using (var reader = OpenXmlReader.Create(workbookStylesPart.Stylesheet))
            {
                while (reader.Read())
                {
                    // numbering formats
                    if (reader.ElementType == typeof(NumberingFormats))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(NumberingFormats) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(NumberingFormat))
                            {
                                this.stylesheet.AddNumberingFormat((NumberingFormat)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }

                    // fonts
                    else if (reader.ElementType == typeof(Fonts))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(Fonts) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(Font))
                            {
                                this.stylesheet.AddFont((Font)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }

                    // fills
                    else if (reader.ElementType == typeof(Fills))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(Fills) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(Fill))
                            {
                                this.stylesheet.AddFill((Fill)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }

                    // borders
                    else if (reader.ElementType == typeof(Borders))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(Borders) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(Border))
                            {
                                this.stylesheet.AddBorder((Border)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }

                    // cell style formats
                    else if (reader.ElementType == typeof(CellStyleFormats))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(CellStyleFormats) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(CellFormat))
                            {
                                this.stylesheet.AddCellStyleFormat((CellFormat)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }

                    // cell formats
                    else if (reader.ElementType == typeof(CellFormats))
                    {
                        reader.ReadFirstChild();
                        if (reader.ElementType == typeof(CellFormats) && reader.IsEndElement)
                        {
                            continue;
                        }

                        do
                        {
                            if (reader.ElementType == typeof(CellFormat))
                            {
                                this.stylesheet.AddCellFormat((CellFormat)reader.LoadCurrentElement());
                            }
                        } while (reader.ReadNextSibling());
                    }
                }
            }
        }

        private void ValidateClassMap<TClass, TClassMap>()
            where TClass : class
            where TClassMap : ClassMap<TClass>
        {
            var classMapType = typeof(TClassMap);
            if (!this.validatedClassMaps.Contains(classMapType))
            {
                var validator = new ConfigurationValidator<TClass, TClassMap>();
                validator.Validate();
                if (validator.HasErrors)
                {
                    throw new AggregateException("One or more property maps have invalid map definitions.", validator.Errors);
                }
            }
            else
            {
                this.validatedClassMaps.Add(classMapType);
            }
        }

        private void WriteSharedStringPart()
        {
            if (this.sharedStrings.Count > this.originalSharedStringCount)
            {
                var sharedStringTablePart = this.spreadsheetDocument.WorkbookPart.SharedStringTablePart;
                if (sharedStringTablePart != null)
                {
                    this.spreadsheetDocument.WorkbookPart.DeletePart(sharedStringTablePart);
                }

                var sharedStringPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                using (var writer = OpenXmlWriter.Create(sharedStringPart))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    foreach (var item in this.sharedStrings)
                    {
                        writer.WriteStartElement(new SharedStringItem());
                        writer.WriteElement(new Text(item.Key));
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                }
            }
        }

        private void WriteStyles()
        {
            var workbookStylesPart = this.spreadsheetDocument.WorkbookPart.WorkbookStylesPart ?? this.spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            workbookStylesPart.Stylesheet = new Stylesheet()
            {
                Borders = new Borders(),
                CellFormats = new CellFormats(),
                CellStyleFormats = new CellStyleFormats(),
                Fills = new Fills(),
                Fonts = new Fonts(),
                NumberingFormats = new NumberingFormats(),
            };

            workbookStylesPart.Stylesheet.Borders.Append(this.stylesheet.Borders.Keys);
            workbookStylesPart.Stylesheet.Fills.Append(this.stylesheet.Fills.Keys);
            workbookStylesPart.Stylesheet.Fonts.Append(this.stylesheet.Fonts.Keys);
            workbookStylesPart.Stylesheet.NumberingFormats.Append(this.stylesheet.NumberingFormats.Keys);

            workbookStylesPart.Stylesheet.CellFormats.Append(this.stylesheet.CellFormats.Keys);
            workbookStylesPart.Stylesheet.CellStyleFormats.Append(this.stylesheet.CellStyleFormats.Keys);

            workbookStylesPart.Stylesheet.Save();
        }

        #region IDisposable

        /// <summary>
        /// Closes the <see cref="Spreadsheet"/> object and the underlying stream, and releases any the system resources associated with the file.
        /// </summary>
        public void Close()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="Spreadsheet"/> object and the underlying stream, and releases any the system resources associated with the file.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Closes the <see cref="Spreadsheet"/> object and the underlying stream, and optionally releases any the system resources associated with the file.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.isDisposed && disposing)
            {
                this.WriteSharedStringPart();
                this.WriteStyles();
                this.spreadsheetDocument.Close();
            }

            this.isDisposed = true;
        }

        #endregion
    }
}