namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;

    using Tests.Models;
    using Xunit;

    public class NumberFormats
    {
        private readonly string filepath;
        private const int recordCount = 25;

        private const string boolWorksheetName = "bool";
        private const string charWorksheetName = "char";
        private const string datetimeWorksheetName = "datetime";
        private const string decimalWorksheetName = "decimal";
        private const string floatWorksheetName = "float";
        private const string longWorksheetName = "long";
        private const string textWorksheetName = "text";

        public NumberFormats()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "number_formats.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapBool>(boolWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapChar>(charWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapDateTime>(datetimeWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapDecimal>(decimalWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapFloat>(floatWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapLong>(longWorksheetName, records);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapText>(textWorksheetName, records);
            }
        }

        [Fact]
        public void TestWrite()
        {
            var validator = new SpreadsheetValidator();
            validator.Validate(this.filepath);

            Assert.False(validator.HasErrors);
        }

        private static IEnumerable<DataTypes> CreateTestRecords(int count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return new DataTypes(true);
            }
        }

        private class TestClassMapBool : ClassMap<DataTypes>
        {
            public TestClassMapBool()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Bool).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapChar : ClassMap<DataTypes>
        {
            public TestClassMapChar()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Char).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapDateTime : ClassMap<DataTypes>
        {
            public TestClassMapDateTime()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.DateTime).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapDecimal : ClassMap<DataTypes>
        {
            public TestClassMapDecimal()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Decimal).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapFloat : ClassMap<DataTypes>
        {
            public TestClassMapFloat()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Float).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapLong : ClassMap<DataTypes>
        {
            public TestClassMapLong()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Long).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }

        private class TestClassMapText : ClassMap<DataTypes>
        {
            public TestClassMapText()
            {
                uint columnIndex = 1;
                foreach (var numberFormat in (OpenXmlNumberingFormat[])Enum.GetValues(typeof(OpenXmlNumberingFormat)))
                {
                    base.Map(x => x.Text).Index(columnIndex).IgnoreRead(true).Name($"{numberFormat.ToString()}").Style(new ColumnStyle() { NumberFormat = numberFormat });
                    columnIndex++;
                }
            }
        }
    }
}