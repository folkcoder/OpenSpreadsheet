namespace Tests.DataConversions
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;
    using Xunit;

    public class NumberFormats : SpreadsheetTesterBase
    {
        private static readonly Dictionary<uint, string> customNumberFormats = new Dictionary<uint, string>()
        {
            { 1, "\"The number is \"#,##0.00" },
            { 2, "\"$#,##0.00;[Black]($#,##0.00);\"" },
            { 3, "\"$#,##0.00;[Red]$#,##0.00\"" },
            { 4, "\"$#,##0.00;[Black]$-#,##0.00;\"" },
        };

        private static readonly Dictionary<uint, OpenXmlNumberingFormat> numberFormats = new Dictionary<uint, OpenXmlNumberingFormat>()
        {
            { 1, OpenXmlNumberingFormat.Accounting },
            { 2, OpenXmlNumberingFormat.Currency },
            { 3, OpenXmlNumberingFormat.DateDayMonth },
            { 4, OpenXmlNumberingFormat.DateDayMonthYear },
            { 5, OpenXmlNumberingFormat.DateMonthDayYear },
            { 6, OpenXmlNumberingFormat.DateMonthYear },
            { 7, OpenXmlNumberingFormat.DatestampWithTime },
            { 8, OpenXmlNumberingFormat.Decimal },
            { 9, OpenXmlNumberingFormat.DecimalWithNegativeInParens },
            { 10, OpenXmlNumberingFormat.DecimalWithNegativeInRedParens },
            { 11, OpenXmlNumberingFormat.FractionOneDigit },
            { 12, OpenXmlNumberingFormat.FractionTwoDigits },
            { 13, OpenXmlNumberingFormat.General },
            { 14, OpenXmlNumberingFormat.Number },
            { 15, OpenXmlNumberingFormat.NumberWithCommas},
            { 16, OpenXmlNumberingFormat.NumberWithNegativeInParens },
            { 17, OpenXmlNumberingFormat.NumberWithNegativeInRedParens },
            { 18, OpenXmlNumberingFormat.Percentage },
            { 19, OpenXmlNumberingFormat.PercentageDecimal},
            { 20, OpenXmlNumberingFormat.Scientific },
            { 21, OpenXmlNumberingFormat.Text },
            { 22, OpenXmlNumberingFormat.TimestampConditionalHourMinuteSecond },
            { 23, OpenXmlNumberingFormat.TimestampHouMinuteSecond12 },
            { 24, OpenXmlNumberingFormat.TimestampHouMinuteSecond24 },
            { 25, OpenXmlNumberingFormat.TimestampHourMinute12 },
            { 26, OpenXmlNumberingFormat.TimestampHourMinute24 },
            { 27, OpenXmlNumberingFormat.TimestampMinuteSecond },
        };

        [Fact]
        public void TestCustomNumberFormat()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapCustomNumberFormats>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedNumberFormat = customNumberFormats[columnIndex];
                        var cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];
                        var numberingFormat = (NumberingFormat)workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.ChildElements.FirstOrDefault(x => ((NumberingFormat)x).NumberFormatId == cellFormat.NumberFormatId);

                        Assert.Equal(expectedNumberFormat, numberingFormat.FormatCode);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestNumberFormats()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapNumberFormats>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedNumberFormat = numberFormats[columnIndex];
                        var cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];

                        Assert.Equal((uint)expectedNumberFormat, (uint)cellFormat.NumberFormatId);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        public class TestClass
        {
            public string Text { get; set; } = "32312.54";
        }

        private class TestClassMapCustomNumberFormats : ClassMap<TestClass>
        {
            public TestClassMapCustomNumberFormats()
            {
                foreach (var numberFormat in customNumberFormats)
                {
                    base.Map(x => x.Text).IgnoreRead(true).Index(numberFormat.Key).Style(new ColumnStyle() { CustomNumberFormat = numberFormat.Value });
                }
            }
        }

        private class TestClassMapNumberFormats : ClassMap<TestClass>
        {
            public TestClassMapNumberFormats()
            {
                foreach (var numberFormat in numberFormats)
                {
                    base.Map(x => x.Text).IgnoreRead(true).Index(numberFormat.Key).Style(new ColumnStyle() { NumberFormat = numberFormat.Value });
                }
            }
        }
    }
}