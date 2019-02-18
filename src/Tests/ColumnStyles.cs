namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;

    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;

    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class ColumnStyles2
    {
        private const string backgroundsSheetName = "Backgrounds";
        private const string borderPlacementsSheetName = "Border Placements";
        private const string borderStylesSheetName = "Border Styles";
        private const int recordCount = 25;
        private readonly string filepath;

        public ColumnStyles2()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "column_styles.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = this.CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBackgroundPatternTypes>(backgroundsSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderPlacement>(borderPlacementsSheetName, records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMapBorderStyles>(borderStylesSheetName, records);
            }
        }

        [Fact]
        public void TestWrite()
        {
            var validator = new SpreadsheetValidator();
            validator.Validate(this.filepath);

            Assert.False(validator.HasErrors);
        }

        private IEnumerable<TestClass> CreateTestRecords(int count)
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

        private class TestClassMapBackgroundPatternTypes : ClassMap<TestClass>
        {
            public TestClassMapBackgroundPatternTypes()
            {
                uint columnIndex = 1;
                foreach (var patternType in (OpenXml.PatternValues[])Enum.GetValues(typeof(OpenXml.PatternValues)))
                {
                    base.Map(x => x.TestData).Index(columnIndex).Name($"{patternType.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BackgroundColor = Color.DarkBlue, BackgroundPatternType = patternType });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapBorderPlacement : ClassMap<TestClass>
        {
            public TestClassMapBorderPlacement()
            {
                uint columnIndex = 1;
                foreach (var borderPlacement in (BorderPlacement[])Enum.GetValues(typeof(BorderPlacement)))
                {
                    base.Map(x => x.TestData).Index(columnIndex).Name($"{borderPlacement.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BorderColor = Color.DarkBlue, BorderPlacement = borderPlacement, BorderStyle = OpenXml.BorderStyleValues.Thick });
                    columnIndex++;
                    base.Map().Constant("").Index(columnIndex).Name("");
                    columnIndex++;
                }
            }
        }

        private class TestClassMapBorderStyles : ClassMap<TestClass>
        {
            public TestClassMapBorderStyles()
            {
                uint columnIndex = 1;
                foreach (var borderStyle in (OpenXml.BorderStyleValues[])Enum.GetValues(typeof(OpenXml.BorderStyleValues)))
                {
                    base.Map(x => x.TestData).Index(columnIndex).Name($"{borderStyle.ToString()}").IgnoreRead(true).Style(new ColumnStyle() { BorderColor = Color.DarkBlue, BorderPlacement = BorderPlacement.All, BorderStyle = borderStyle });
                    columnIndex++;
                    base.Map().Constant("").Index(columnIndex).Name("");
                    columnIndex++;
                }
            }
        }
    }
}