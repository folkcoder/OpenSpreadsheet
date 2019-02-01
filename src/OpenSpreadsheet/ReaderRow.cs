namespace OpenSpreadsheet
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Encapsulates properties associated with a spreadsheet row.
    /// </summary>
    public class ReaderRow
    {
        private readonly BidirectionalDictionary<uint, string> headers;
        private readonly Dictionary<uint, string> rowValues;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReaderRow"/> class.
        /// </summary>
        /// <param name="headers">A dictionary containing header column indexes and their associated names.</param>
        /// <param name="rowValues">A dictionary containing row column indeces and their associated values.</param>
        public ReaderRow(BidirectionalDictionary<uint, string> headers, Dictionary<uint, string> rowValues)
        {
            this.headers = headers;
            this.rowValues = rowValues;
        }

        /// <summary>
        /// Retrieves the value of the cell at the provided column.
        /// </summary>
        /// <param name="headerName">The column header name.</param>
        /// <returns>The cell's value.</returns>
        public string GetCellValue(string headerName)
        {
            this.headers.TryGetKey(headerName, out uint columnIndex);
            if (columnIndex == 0)
            {
                throw new KeyNotFoundException();
            }

            this.rowValues.TryGetValue(columnIndex, out string value);
            if (value == null)
            {
                throw new KeyNotFoundException();
            }

            return value;
        }

        /// <summary>
        /// Retrieves the value of the cell at the specified column index.
        /// </summary>
        /// <param name="columnIndex">The one-based column index where the cell is located.</param>
        /// <returns>The cell's value.</returns>
        public string GetCellValue(uint columnIndex)
        {
            this.rowValues.TryGetValue(columnIndex, out string value);
            if (value == null)
            {
                throw new KeyNotFoundException();
            }

            return value;
        }
    }
}