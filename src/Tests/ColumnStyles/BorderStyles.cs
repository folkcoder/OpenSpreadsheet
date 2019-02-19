namespace Tests.ColumnStyles
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

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class BorderStyles : SpreadsheetTesterBase
    {
        private static readonly Dictionary<uint, OpenXml.BorderStyleValues> borderColors = new Dictionary<uint, OpenXml.BorderStyleValues>()
        {
            { 1, OpenXml.BorderStyleValues.DashDot },
            { 2, OpenXml.BorderStyleValues.DashDotDot },
            { 3, OpenXml.BorderStyleValues.Dashed },
            { 4, OpenXml.BorderStyleValues.Dotted },
            { 5, OpenXml.BorderStyleValues.Double },
            { 6, OpenXml.BorderStyleValues.Hair },
            { 7, OpenXml.BorderStyleValues.Medium },
            { 8, OpenXml.BorderStyleValues.MediumDashDot },
            { 9, OpenXml.BorderStyleValues.MediumDashDotDot },
            { 10, OpenXml.BorderStyleValues.MediumDashed },
            { 11, OpenXml.BorderStyleValues.None },
            { 12, OpenXml.BorderStyleValues.SlantDashDot },
            { 13, OpenXml.BorderStyleValues.Thick},
            { 14, OpenXml.BorderStyleValues.Thin},
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
            { 1, OpenXml.BorderStyleValues.DashDot },
            { 2, OpenXml.BorderStyleValues.DashDotDot },
            { 3, OpenXml.BorderStyleValues.Dashed },
            { 4, OpenXml.BorderStyleValues.Dotted },
            { 5, OpenXml.BorderStyleValues.Double },
            { 6, OpenXml.BorderStyleValues.Hair },
            { 7, OpenXml.BorderStyleValues.Medium },
            { 8, OpenXml.BorderStyleValues.MediumDashDot },
            { 9, OpenXml.BorderStyleValues.MediumDashDotDot },
            { 10, OpenXml.BorderStyleValues.MediumDashed },
            { 11, OpenXml.BorderStyleValues.None },
            { 12, OpenXml.BorderStyleValues.SlantDashDot },
            { 13, OpenXml.BorderStyleValues.Thick},
            { 14, OpenXml.BorderStyleValues.Thin},
        };

        [Fact]
        public void TestBorderStyles()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderStyles>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(this.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedBorderStyle = borderStyles[columnIndex];
                        var border = (Border)workbookPart.WorkbookStylesPart.Stylesheet.Borders.ChildElements[(int)cell.StyleIndex.Value];

                        Assert.Equal<BorderStyleValues>(expectedBorderStyle, border.BottomBorder.Style);
                        Assert.Equal<BorderStyleValues>(expectedBorderStyle, border.DiagonalBorder.Style);
                        Assert.Equal<BorderStyleValues>(expectedBorderStyle, border.LeftBorder.Style);
                        Assert.Equal<BorderStyleValues>(expectedBorderStyle, border.RightBorder.Style);
                        Assert.Equal<BorderStyleValues>(expectedBorderStyle, border.TopBorder.Style);
                    }
                }
            }

            File.Delete(filepath);
        }

        [Fact]
        public void TestBorderPlacements()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderPlacements>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(this.SpreadsheetValidator.HasErrors);

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedBorderPlacement = borderPlacements[columnIndex];
                        var border = (Border)workbookPart.WorkbookStylesPart.Stylesheet.Borders.ChildElements[(int)cell.StyleIndex.Value];

                        if (border.BottomBorder != null)
                        {
                            Assert.True(expectedBorderPlacement.HasFlag(BorderPlacement.Bottom));
                        }

                        if (border.DiagonalBorder != null)
                        {
                            Assert.Equal(expectedBorderPlacement.HasFlag(BorderPlacement.DiagonalDown), border.DiagonalDown.Value);
                            Assert.Equal(expectedBorderPlacement.HasFlag(BorderPlacement.DiagonalUp), border.DiagonalUp.Value);
                        }

                        if (border.LeftBorder != null)
                        {
                            Assert.True(expectedBorderPlacement.HasFlag(BorderPlacement.Left));
                        }

                        if (border.RightBorder != null)
                        {
                            Assert.True(expectedBorderPlacement.HasFlag(BorderPlacement.Right));
                        }

                        if (border.TopBorder != null)
                        {
                            Assert.True(expectedBorderPlacement.HasFlag(BorderPlacement.Top));
                        }
                    }
                }
            }

            File.Delete(filepath);
        }

        private class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        private class TestClassMapBorderPlacements : ClassMap<TestClass>
        {
            public TestClassMapBorderPlacements()
            {
                foreach (var borderPlacement in borderPlacements)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(borderPlacement.Key).Style(new ColumnStyle() { BorderPlacement = borderPlacement.Value, BorderStyle = BorderStyleValues.Thick });
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