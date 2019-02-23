namespace Tests.Styles
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class Fonts : SpreadsheetTesterBase
    {
        private static readonly Dictionary<uint, Font> fonts = new Dictionary<uint, Font>()
        {
            { 1, new Font("Garamond", 12, FontStyle.Regular) },
            { 2, new Font("Garamond", 12, FontStyle.Underline) },
            { 3, new Font("Times New Roman", 12, FontStyle.Italic) },
            { 4, new Font("Wingdings", 14, FontStyle.Strikeout) },
            { 5, new Font( FontFamily.GenericMonospace, 35, FontStyle.Bold) },
            { 6, new Font( FontFamily.Families[10], 4, FontStyle.Bold | FontStyle.Underline | FontStyle.Italic ) },
        };

        private static readonly Dictionary<uint, Color> foregroundColors = new Dictionary<uint, Color>()
        {
            { 1, Color.Chocolate },
            { 2, Color.Teal },
            { 3, Color.Black },
            { 4, Color.White },
            { 5, Color.BurlyWood }
        };

        [Fact]
        public void TestFonts()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapFonts>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(base.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheet = worksheetPart.Worksheet;

                foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                {
                    var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                    var expectedFont = fonts[columnIndex];
                    var cellFont = (OpenXml.Font)workbookPart.WorkbookStylesPart.Stylesheet.Fonts.ChildElements[(int)cell.StyleIndex.Value];

                    if (cellFont.Bold != null)
                    {
                        Assert.True(expectedFont.Bold);
                    }

                    if (cellFont.Italic != null)
                    {
                        Assert.True(expectedFont.Italic);
                    }

                    if (cellFont.Strike != null)
                    {
                        Assert.True(expectedFont.Strikeout);
                    }

                    if (cellFont.Underline != null)
                    {
                        Assert.True(expectedFont.Underline);
                    }

                    Assert.Equal(expectedFont.FontFamily.Name, cellFont.FontName.Val);
                    Assert.Equal(expectedFont.Size, cellFont.FontSize.Val);
                }
            }

            File.Delete(filepath);
        }

        [Fact]
        public void TestForegrounds()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapForegrounds>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(base.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheet = worksheetPart.Worksheet;

                foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                {
                    var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                    var expectedForeground = foregroundColors[columnIndex];
                    var cellFont = (OpenXml.Font)workbookPart.WorkbookStylesPart.Stylesheet.Fonts.ChildElements[(int)cell.StyleIndex.Value];

                    Assert.Equal(base.ConvertColorToHex(expectedForeground), cellFont.Color.Rgb.Value);
                }
            }

            File.Delete(filepath);
        }

        private class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        private class TestClassMapFonts : ClassMap<TestClass>
        {
            public TestClassMapFonts()
            {
                foreach (var font in fonts)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(font.Key).Style(new ColumnStyle() { Font = font.Value });
                }
            }
        }

        private class TestClassMapForegrounds : ClassMap<TestClass>
        {
            public TestClassMapForegrounds()
            {
                foreach (var foregroundColor in foregroundColors)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(foregroundColor.Key).Style(new ColumnStyle() { ForegroundColor = foregroundColor.Value });
                }
            }
        }
    }
}