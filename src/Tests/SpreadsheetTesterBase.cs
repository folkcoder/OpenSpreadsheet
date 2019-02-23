namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using DocumentFormat.OpenXml.Packaging;
    using Excel = Microsoft.Office.Interop.Excel;

    public abstract class SpreadsheetTesterBase
    {
        protected SpreadsheetValidator SpreadsheetValidator { get; } = new SpreadsheetValidator();

        public string ConstructTempXlsxSaveName() => Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".xlsx"));

        public uint GetColumnIndexFromCellReference(string cellReference)
        {
            var columnLetters = cellReference.Where(c => !char.IsNumber(c)).ToArray();
            string columnReference = new string(columnLetters);

            int sum = 0;
            for (int i = 0; i < columnLetters.Length; i++)
            {
                sum *= 26;
                sum += columnLetters[i] - 'A' + 1;
            }

            return (uint)sum;
        }

        public string ConvertColorToHex(in Color color) => Color.FromArgb(color.ToArgb()).Name;

        public IEnumerable<T> CreateRecords<T>(int count) where T : class
        {
            for (int i = 0; i < count; i++)
            {
                yield return Activator.CreateInstance<T>();
            }
        }

        public string GetSharedStringValue(SharedStringTablePart sharedStringTablePart, string cellValue) =>
            sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;

        public string SaveAsExcelFile(string spreadsheetFile)
        {
            var excelFileName = this.ConstructTempXlsxSaveName();
            var excel = new Excel.Application();
            var workbook = excel.Workbooks.Open(spreadsheetFile);
            workbook.SaveAs(excelFileName);
            workbook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);

            return excelFileName;
        }
    }
}