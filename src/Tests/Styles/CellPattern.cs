namespace Tests.Styles
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class CellPattern : SpreadsheetTesterBase
    {
        // Fill index begins at 2, following default system fills.
        private static readonly Dictionary<uint, Color> backgrounds = new Dictionary<uint, Color>()
        {
            { 2, Color.Chocolate },
            { 3, Color.Teal },
            { 4, Color.Black },
            { 5, Color.White },
            { 6, Color.BurlyWood },
        };

        // skips OpenXml.PatternValues.Gray125 }, system pattern at index 1
        private static readonly Dictionary<uint, OpenXml.PatternValues> patterns = new Dictionary<uint, OpenXml.PatternValues>()
        {
            { 2, OpenXml.PatternValues.DarkDown },
            { 3, OpenXml.PatternValues.DarkGray  },
            { 4, OpenXml.PatternValues.DarkGrid },
            { 5, OpenXml.PatternValues.DarkHorizontal },
            { 6, OpenXml.PatternValues.DarkTrellis },
            { 7, OpenXml.PatternValues.DarkUp },
            { 8, OpenXml.PatternValues.DarkVertical },
            { 9, OpenXml.PatternValues.Gray0625 },
            { 10, OpenXml.PatternValues.LightDown },
            { 11, OpenXml.PatternValues.LightGray },
            { 12, OpenXml.PatternValues.LightGrid },
            { 13, OpenXml.PatternValues.LightHorizontal },
            { 14, OpenXml.PatternValues.LightTrellis },
            { 15, OpenXml.PatternValues.LightUp },
            { 16, OpenXml.PatternValues.LightVertical },
            { 17, OpenXml.PatternValues.MediumGray },
            { 18, OpenXml.PatternValues.None },
            { 19, OpenXml.PatternValues.Solid },
        };

        [Fact]
        public void TestBackgrounds()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBackgrounds>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(this.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheet = worksheetPart.Worksheet;

                foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                {
                    var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                    var expectedColor = backgrounds[columnIndex];
                    var fill = (OpenXml.Fill)workbookPart.WorkbookStylesPart.Stylesheet.Fills.ChildElements[(int)columnIndex];

                    Assert.Equal(base.ConvertColorToHex(expectedColor), fill.PatternFill.ForegroundColor.Rgb.Value);
                }
            }

            File.Delete(filepath);
        }

        [Fact]
        public void TestPatterns()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapPatterns>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(this.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheet = worksheetPart.Worksheet;

                foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                {
                    var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                    var expectedPattern = patterns[columnIndex];
                    var fill = (OpenXml.Fill)workbookPart.WorkbookStylesPart.Stylesheet.Fills.ChildElements[(int)columnIndex];

                    Assert.Equal<OpenXml.PatternValues>(expectedPattern, fill.PatternFill.PatternType);
                }
            }

            File.Delete(filepath);
        }

        private class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        private class TestClassMapBackgrounds: ClassMap<TestClass>
        {
            public TestClassMapBackgrounds()
            {
                foreach (var color in backgrounds)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(color.Key).Style(new ColumnStyle() { BackgroundColor = color.Value });
                }
            }
        }

        private class TestClassMapPatterns : ClassMap<TestClass>
        {
            public TestClassMapPatterns()
            {
                foreach (var pattern in patterns)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(pattern.Key).Style(new ColumnStyle() { BackgroundColor = Color.DarkBlue, BackgroundPatternType = pattern.Value });
                }
            }
        }
    }
}