namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using OpenSpreadsheet.Enums;
    using Xunit;

    public class CustomNumberFormats
    {
        private const decimal currencyValue = -52432.23M;
        private const int recordCount = 25;
        private const string customNumberFormatsWorksheetName = "custom formats";

        private readonly string filepath;

        public CustomNumberFormats()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "custom_number_formats.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>(customNumberFormatsWorksheetName, records);
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

        internal class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.DefaultCurrencyFormat).Index(1).Style(new ColumnStyle() { NumberFormat = OpenXmlNumberingFormat.Currency });
                base.Map(x => x.NegativeInParens).Index(2).Style(new ColumnStyle() { CustomNumberFormat = "$#,##0.00;[Black]($#,##0.00);" });
                base.Map(x => x.NegativeInRed).Index(4).Style(new ColumnStyle() { CustomNumberFormat = "$#,##0.00;[Red]$#,##0.00;" });
                base.Map(x => x.NegativeWithSign).Index(3).Style(new ColumnStyle() { CustomNumberFormat = "$#,##0.00;[Black]$-#,##0.00;" });
            }
        }

        internal class TestClass
        {
            public decimal DefaultCurrencyFormat { get; set; } = currencyValue;
            public decimal NegativeInParens { get; set; } = currencyValue;
            public decimal NegativeInRed { get; set; } = currencyValue;
            public decimal NegativeWithSign { get; set; } = currencyValue;
        }
    }
}