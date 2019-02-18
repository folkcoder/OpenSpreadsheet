namespace Tests.ColumnStyles
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class Alignments
    {
        private readonly Dictionary<uint, HorizontalAlignmentValues> horizontalAlignments = new Dictionary<uint, HorizontalAlignmentValues>()
        {
            { 1, HorizontalAlignmentValues.Center },
            { 2, HorizontalAlignmentValues.CenterContinuous },
            { 3, HorizontalAlignmentValues.Distributed },
            { 4, HorizontalAlignmentValues.Fill },
            { 5, HorizontalAlignmentValues.General },
            { 6, HorizontalAlignmentValues.Justify },
            { 7, HorizontalAlignmentValues.Left },
            { 8, HorizontalAlignmentValues.Right },
        };

        private readonly Dictionary<uint, VerticalAlignmentValues> verticalAlignments = new Dictionary<uint, VerticalAlignmentValues>()
        {
            { 1, VerticalAlignmentValues.Bottom },
            { 2, VerticalAlignmentValues.Center },
            { 3, VerticalAlignmentValues.Distributed },
            { 4, VerticalAlignmentValues.Justify },
            { 5, VerticalAlignmentValues.Top },
        };

        [Fact]
        public void TestHorizontalAlignments()
        {
            var filepath = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapHorizontalAlignments>("Sheet1", CreateTestRecords(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = SpreadsheetHelpers.GetColumnIndexFromCellReference(cell.CellReference);
                        this.horizontalAlignments.TryGetValue(columnIndex, out HorizontalAlignmentValues expectedAlignment);

                        var cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];

                        Assert.Equal<HorizontalAlignmentValues>(expectedAlignment, cellFormat.Alignment.Horizontal);
                    }
                }
            }

            File.Delete(filepath);
        }

        [Fact]
        public void TestVerticalAlignments()
        {
            var filepath = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapVerticalAlignments>("Sheet1", CreateTestRecords(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }

            using (var filestream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = SpreadsheetHelpers.GetColumnIndexFromCellReference(cell.CellReference);
                        this.verticalAlignments.TryGetValue(columnIndex, out VerticalAlignmentValues expectedAlignment);

                        var cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];

                        Assert.Equal<VerticalAlignmentValues>(expectedAlignment, cellFormat.Alignment.Vertical);
                    }
                }
            }

            File.Delete(filepath);
        }

        private static IEnumerable<TestClass> CreateTestRecords(int count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return new TestClass();
            }
        }

        private class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        private class TestClassMapHorizontalAlignments : ClassMap<TestClass>
        {
            public TestClassMapHorizontalAlignments()
            {
                base.Map(x => x.TestData).Index(1).IgnoreRead(true).Name("Center").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Center });
                base.Map(x => x.TestData).Index(2).IgnoreRead(true).Name("CenterContinuous").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.CenterContinuous });
                base.Map(x => x.TestData).Index(3).IgnoreRead(true).Name("Distributed").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Distributed });
                base.Map(x => x.TestData).Index(4).IgnoreRead(true).Name("Fill").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Fill });
                base.Map(x => x.TestData).Index(5).IgnoreRead(true).Name("General").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.General });
                base.Map(x => x.TestData).Index(6).IgnoreRead(true).Name("Justify").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Justify });
                base.Map(x => x.TestData).Index(7).IgnoreRead(true).Name("Left").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Left });
                base.Map(x => x.TestData).Index(8).IgnoreRead(true).Name("Right").Style(new ColumnStyle() { HoizontalAlignment = HorizontalAlignmentValues.Right });
            }
        }

        private class TestClassMapVerticalAlignments : ClassMap<TestClass>
        {
            public TestClassMapVerticalAlignments()
            {
                base.Map(x => x.TestData).Index(1).IgnoreRead(true).Name("Bottom").Style(new ColumnStyle() { VerticalAlignment = VerticalAlignmentValues.Bottom });
                base.Map(x => x.TestData).Index(2).IgnoreRead(true).Name("Center").Style(new ColumnStyle() { VerticalAlignment = VerticalAlignmentValues.Center });
                base.Map(x => x.TestData).Index(3).IgnoreRead(true).Name("Distributed").Style(new ColumnStyle() { VerticalAlignment = VerticalAlignmentValues.Distributed });
                base.Map(x => x.TestData).Index(4).IgnoreRead(true).Name("Justify").Style(new ColumnStyle() { VerticalAlignment = VerticalAlignmentValues.Justify });
                base.Map(x => x.TestData).Index(5).IgnoreRead(true).Name("Top").Style(new ColumnStyle() { VerticalAlignment = VerticalAlignmentValues.Top });
            }
        }
    }
}