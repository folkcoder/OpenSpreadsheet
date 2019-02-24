namespace Tests.Styles
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;
    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class BorderStyles : SpreadsheetTesterBase
    {
        private static readonly Dictionary<uint, Color> borderColors = new Dictionary<uint, Color>()
        {
            { 1, Color.Chocolate },
            { 2, Color.Teal },
            { 3, Color.Black },
            { 4, Color.White },
            { 5, Color.BurlyWood }
        };

        private static readonly Dictionary<uint, BorderPlacement> borderPlacements = new Dictionary<uint, BorderPlacement>()
        {
            { 1, BorderPlacement.All },
            { 2, BorderPlacement.Bottom },
            { 3, BorderPlacement.DiagonalDown },
            { 4, BorderPlacement.DiagonalUp },
            { 5, BorderPlacement.Left},
            { 6, BorderPlacement.Outside },
            { 7, BorderPlacement.Right },
            { 8, BorderPlacement.Top },
            { 9, BorderPlacement.Bottom | BorderPlacement.Top },
            { 10, BorderPlacement.DiagonalDown | BorderPlacement.DiagonalUp },
        };

        private static readonly Dictionary<uint, OpenXml.BorderStyleValues> borderStyles = new Dictionary<uint, OpenXml.BorderStyleValues>()
        {
            { 1, OpenXml.BorderStyleValues.None },
            { 2, OpenXml.BorderStyleValues.DashDot },
            { 3, OpenXml.BorderStyleValues.DashDotDot },
            { 4, OpenXml.BorderStyleValues.Dashed },
            { 5, OpenXml.BorderStyleValues.Dotted },
            { 6, OpenXml.BorderStyleValues.Double },
            { 7, OpenXml.BorderStyleValues.Hair },
            { 8, OpenXml.BorderStyleValues.Medium },
            { 9, OpenXml.BorderStyleValues.MediumDashDot },
            { 10, OpenXml.BorderStyleValues.MediumDashDotDot },
            { 11, OpenXml.BorderStyleValues.MediumDashed },
            { 12, OpenXml.BorderStyleValues.SlantDashDot },
            { 13, OpenXml.BorderStyleValues.Thick},
            { 14, OpenXml.BorderStyleValues.Thin},
        };

        [Fact]
        public void TestBorderColors()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderColors>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
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

                    foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedColor = base.ConvertColorToHex(borderColors[columnIndex]);
                        var border = (OpenXml.Border)workbookPart.WorkbookStylesPart.Stylesheet.Borders.ChildElements[(int)columnIndex];

                        Assert.Equal(expectedColor, border.BottomBorder.Color.Rgb.Value, true);
                        Assert.Equal(expectedColor, border.DiagonalBorder.Color.Rgb.Value, true);
                        Assert.Equal(expectedColor, border.LeftBorder.Color.Rgb.Value, true);
                        Assert.Equal(expectedColor, border.RightBorder.Color.Rgb.Value, true);
                        Assert.Equal(expectedColor, border.TopBorder.Color.Rgb.Value, true);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestBorderPlacements()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderPlacements>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
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

                    foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedBorderPlacement = borderPlacements[columnIndex];
                        var border = (OpenXml.Border)workbookPart.WorkbookStylesPart.Stylesheet.Borders.ChildElements[(int)columnIndex];

                        if (border.BottomBorder?.HasChildren == true)
                        {
                            Assert.True((expectedBorderPlacement & BorderPlacement.Bottom) != 0);
                        }

                        if (border.DiagonalDown?.HasValue == true)
                        {
                            Assert.Equal((expectedBorderPlacement & BorderPlacement.DiagonalDown) != 0, border.DiagonalDown.Value);
                        }

                        if (border.DiagonalUp?.HasValue == true)
                        {
                            Assert.Equal((expectedBorderPlacement & BorderPlacement.DiagonalUp) != 0, border.DiagonalUp.Value);
                        }

                        if (border.LeftBorder?.HasChildren == true)
                        {
                            Assert.True((expectedBorderPlacement & BorderPlacement.Left) != 0);
                        }

                        if (border.RightBorder?.HasChildren == true)
                        {
                            Assert.True((expectedBorderPlacement & BorderPlacement.Right) != 0);
                        }

                        if (border.TopBorder?.HasChildren == true)
                        {
                            Assert.True((expectedBorderPlacement & BorderPlacement.Top) != 0);
                        }
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestBorderStyles()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderStyles>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
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

                    foreach (var cell in sheet.Descendants<OpenXml.Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedBorderStyle = borderStyles[columnIndex];
                        var border = (OpenXml.Border)workbookPart.WorkbookStylesPart.Stylesheet.Borders.ChildElements[(int)columnIndex];

                        if (expectedBorderStyle == OpenXml.BorderStyleValues.None)
                        {
                            Assert.True(border.BottomBorder.Style == null || border.BottomBorder.Style == OpenXml.BorderStyleValues.None);
                            Assert.True(border.DiagonalBorder.Style == null || border.BottomBorder.Style == OpenXml.BorderStyleValues.None);
                            Assert.True(border.LeftBorder.Style == null || border.BottomBorder.Style == OpenXml.BorderStyleValues.None);
                            Assert.True(border.RightBorder.Style == null || border.BottomBorder.Style == OpenXml.BorderStyleValues.None);
                            Assert.True(border.TopBorder.Style == null || border.BottomBorder.Style == OpenXml.BorderStyleValues.None);
                        }
                        else
                        {
                            Assert.Equal<OpenXml.BorderStyleValues>(expectedBorderStyle, border.BottomBorder.Style);
                            Assert.Equal<OpenXml.BorderStyleValues>(expectedBorderStyle, border.DiagonalBorder.Style);
                            Assert.Equal<OpenXml.BorderStyleValues>(expectedBorderStyle, border.LeftBorder.Style);
                            Assert.Equal<OpenXml.BorderStyleValues>(expectedBorderStyle, border.RightBorder.Style);
                            Assert.Equal<OpenXml.BorderStyleValues>(expectedBorderStyle, border.TopBorder.Style);
                        }
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        private class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        private class TestClassMapBorderColors : ClassMap<TestClass>
        {
            public TestClassMapBorderColors()
            {
                foreach (var color in borderColors)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(color.Key).Style(new ColumnStyle() { BorderColor =  color.Value, BorderPlacement = BorderPlacement.All, BorderStyle = OpenXml.BorderStyleValues.Thick });
                }
            }
        }

        private class TestClassMapBorderPlacements : ClassMap<TestClass>
        {
            public TestClassMapBorderPlacements()
            {
                foreach (var borderPlacement in borderPlacements)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(borderPlacement.Key).Style(new ColumnStyle() { BorderPlacement = borderPlacement.Value, BorderStyle = OpenXml.BorderStyleValues.Thick });
                }
            }
        }

        private class TestClassMapBorderStyles : ClassMap<TestClass>
        {
            public TestClassMapBorderStyles()
            {
                foreach (var borderStyle in borderStyles)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(borderStyle.Key).Style(new ColumnStyle() { BorderPlacement = OpenSpreadsheet.Enums.BorderPlacement.All, BorderStyle = borderStyle.Value });
                }
            }
        }
    }
}