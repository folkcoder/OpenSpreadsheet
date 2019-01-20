namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;

    using SpreadsheetHelper;
    using SpreadsheetHelper.Configuration;
    using SpreadsheetHelper.Enums;

    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class ColumnStyles
    {
        private const string backgroundsSheetName = "Backgrounds";
        private const string borderPlacementsSheetName = "Border Placements";
        private const string borderStylesSheetName = "Border Styles";
        private const string fontsSheetName = "Fonts";
        private const string horizontalAlignmentsSheetName = "Horizontal Alignments";
        private const int recordCount = 25;
        private const string verticalAlignmentsSheetName = "Vertical Alignments";
        private readonly string filepath;

        public ColumnStyles()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "excel tests");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "column_styles.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBackgroundPatternTypes>(backgroundsSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderPlacement>(borderPlacementsSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderStyles>(borderStylesSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapFonts>(fontsSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapHorizontalAlignments>(horizontalAlignmentsSheetName, records, new WorksheetStyle() { MinColumnWidth = 25 });
                spreadsheet.WriteWorksheet<TestClass, TestClassMapVerticalAlignments>(verticalAlignmentsSheetName, records);
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
            for (int i = 0; i < count; i++)
            {
                yield return new TestClass();
            }
        }

        internal class TestClass
        {
            public string TestData { get; set; } = "test data";
        }

        internal class TestClassMapBackgroundPatternTypes : ClassMap<TestClass>
        {
            public TestClassMapBackgroundPatternTypes()
            {
                uint columnIndex = 1;
                foreach (var patternType in (OpenXml.PatternValues[])Enum.GetValues(typeof(OpenXml.PatternValues)))
                {
                    Map(x => x.TestData).Index(columnIndex).Name($"{patternType.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BackgroundColor = Color.DarkBlue, BackgroundPatternType = patternType });
                    columnIndex++;
                }
            }
        }

        internal class TestClassMapBorderPlacement : ClassMap<TestClass>
        {
            public TestClassMapBorderPlacement()
            {
                uint columnIndex = 1;
                foreach (var borderPlacement in (BorderPlacement[])Enum.GetValues(typeof(BorderPlacement)))
                {
                    Map(x => x.TestData).Index(columnIndex).Name($"{borderPlacement.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BorderColor = Color.DarkBlue, BorderPlacement = borderPlacement, BorderStyle = OpenXml.BorderStyleValues.Thick });
                    columnIndex++;
                    Map().Constant("").Index(columnIndex).Name("");
                    columnIndex++;
                }
            }
        }

        internal class TestClassMapBorderStyles : ClassMap<TestClass>
        {
            public TestClassMapBorderStyles()
            {
                uint columnIndex = 1;
                foreach (var borderStyle in (OpenXml.BorderStyleValues[])Enum.GetValues(typeof(OpenXml.BorderStyleValues)))
                {
                    Map(x => x.TestData).Index(columnIndex).Name($"{borderStyle.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BorderColor = Color.DarkBlue, BorderPlacement = BorderPlacement.All, BorderStyle = borderStyle });
                    columnIndex++;
                    Map().Constant("").Index(columnIndex).Name("");
                    columnIndex++;
                }
            }
        }

        internal class TestClassMapFonts : ClassMap<TestClass>
        {
            public TestClassMapFonts()
            {
                Map(x => x.TestData).Index(1).Name("Garamond, 12, Underline, Green").IgnoreRead(true).Style(new ColumnStyle() { Font = new Font("Garamond", 12, FontStyle.Underline), ForegroundColor = Color.Green });
                Map(x => x.TestData).Index(2).Name("TNR, 12, Italic, Default").IgnoreRead(true).Style(new ColumnStyle() { Font = new Font("Times New Roman", 12, FontStyle.Italic) });
                Map(x => x.TestData).Index(3).Name("Wingdings, 14, Strikeout, Tomato").IgnoreRead(true).Style(new ColumnStyle() { Font = new Font("Wingdings", 14, FontStyle.Strikeout), ForegroundColor = Color.Tomato });
            }
        }

        internal class TestClassMapHorizontalAlignments : ClassMap<TestClass>
        {
            public TestClassMapHorizontalAlignments()
            {
                uint columnIndex = 1;
                foreach (var alignment in (OpenXml.HorizontalAlignmentValues[])Enum.GetValues(typeof(OpenXml.HorizontalAlignmentValues)))
                {
                    Map(x => x.TestData).Index(columnIndex).Name($"{alignment.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { HoizontalAlignment = alignment });
                    columnIndex++;
                }
            }
        }

        internal class TestClassMapVerticalAlignments : ClassMap<TestClass>
        {
            public TestClassMapVerticalAlignments()
            {
                uint columnIndex = 1;
                foreach (var alignment in (OpenXml.VerticalAlignmentValues[])Enum.GetValues(typeof(OpenXml.VerticalAlignmentValues)))
                {
                    Map(x => x.TestData).Index(columnIndex).Name($"{alignment.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { VerticalAlignment = alignment });
                    columnIndex++;
                }
            }
        }
    }
}