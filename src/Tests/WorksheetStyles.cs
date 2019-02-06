namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;

    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class WorksheetStyles
    {
        private readonly string filepath;

        public WorksheetStyles()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "worksheet_styles.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(25);

                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Default Style", records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Chartreuse Header Background", records, new WorksheetStyle() { HeaderBackgroundColor = Color.Chartreuse, HeaderBackgroundPatternType = OpenXml.PatternValues.Solid });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Comic Sans 16 Italic Font", records, new WorksheetStyle() { HeaderFont = new Font("Comic Sans", 16, FontStyle.Italic) });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Dark Green Header Foreground", records, new WorksheetStyle() { HeaderForegroundColor = Color.DarkGreen });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Centered Horizontal Headers", records, new WorksheetStyle() { HeaderHoizontalAlignment = OpenXml.HorizontalAlignmentValues.Center });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Top Vertical Headers", records, new WorksheetStyle() { HeaderVerticalAlignment = OpenXml.VerticalAlignmentValues.Top });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Min Col Width 10", records, new WorksheetStyle() { MinColumnWidth = 10 });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Max Col Width 5", records, new WorksheetStyle() { MaxColumnWidth = 5 });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Autofilter On", records, new WorksheetStyle() { ShouldAutoFilter = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Freeze Top Row", records, new WorksheetStyle() { ShouldFreezeTopRow = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("AutoFit Columns", records, new WorksheetStyle() { ShouldAutoFitColumns = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Header Not Written", records, new WorksheetStyle() { ShouldWriteHeaderRow = false });
            }
        }

        [Fact]
        public void TestWrite()
        {
            var validator = new SpreadsheetValidator();
            validator.Validate(this.filepath);

            Assert.False(validator.HasErrors);
        }

        private static IEnumerable<TestClass> CreateTestRecords(int count)
        {
            for (var i = 0; i < count; i++)
            {
                yield return new TestClass();
            }
        }

        private class TestClass
        {
            public string LongText { get; set; } = "The quick brown fox jumps over the lazy dog.";

            public string ShortText { get; set; } = "B";
        }

        private class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.LongText).Index(1).Name("1");
                base.Map(x => x.ShortText).Index(2).Name("2");
            }
        }
    }
}