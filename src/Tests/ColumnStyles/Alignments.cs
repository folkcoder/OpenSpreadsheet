namespace Tests.ColumnStyles
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class Alignments
    {
        private static readonly Dictionary<uint, HorizontalAlignmentValues> horizontalAlignments = new Dictionary<uint, HorizontalAlignmentValues>()
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

        private static readonly Dictionary<uint, VerticalAlignmentValues> verticalAlignments = new Dictionary<uint, VerticalAlignmentValues>()
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
                        var expectedAlignment = horizontalAlignments[columnIndex];

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
                        verticalAlignments.TryGetValue(columnIndex, out VerticalAlignmentValues expectedAlignment);

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
                yield return new TestClass() { TestData = "test data", };
            }
        }

        private class TestClass
        {
            public string TestData { get; set; }
        }

        private class TestClassMapHorizontalAlignments : ClassMap<TestClass>
        {
            public TestClassMapHorizontalAlignments()
            {
                foreach (var alignment in horizontalAlignments)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(alignment.Key).Style(new ColumnStyle() { HoizontalAlignment = alignment.Value });
                }
            }
        }

        private class TestClassMapVerticalAlignments : ClassMap<TestClass>
        {
            public TestClassMapVerticalAlignments()
            {
                foreach (var alignment in verticalAlignments)
                {
                    base.Map(x => x.TestData).IgnoreRead(true).Index(alignment.Key).Style(new ColumnStyle() { VerticalAlignment = alignment.Value });
                }
            }
        }
    }
}