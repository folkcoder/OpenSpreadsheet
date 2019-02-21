namespace Tests.DataConversions
{
    using System;
    using System.IO;

    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    using Xunit;

    public class Defaults : SpreadsheetTesterBase
    {
        private static readonly TestClass testClassWithDefaultValues = new TestClass()
        {
            Bool = true,
            Byte = 24,
            Char = 'h',
            DateTime = new DateTime(2019, 10, 31),
            Decimal = 433123.12M,
            Double = 6.1E+3,
            Float = 23342.93F,
            Int = 551412,
            Long = 4324823423423,
            Text = "default string",
        };

        [Fact]
        public void TestRead()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(base.SpreadsheetValidator.HasErrors);

            using (var spreadsheet = new Spreadsheet(filepath))
            {
                foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMapDefaults>("Sheet1"))
                {
                    Assert.Equal(testClassWithDefaultValues.Bool, record.Bool.Value);
                    Assert.Equal(testClassWithDefaultValues.Byte, record.Byte.Value);
                    Assert.Equal(testClassWithDefaultValues.Char, record.Char.Value);
                    Assert.Equal(testClassWithDefaultValues.DateTime, record.DateTime.Value);
                    Assert.Equal(testClassWithDefaultValues.Decimal, record.Decimal.Value);
                    Assert.Equal(testClassWithDefaultValues.Double, record.Double.Value);
                    Assert.Equal(testClassWithDefaultValues.Float, record.Float.Value);
                    Assert.Equal(testClassWithDefaultValues.Int, record.Int.Value);
                    Assert.Equal(testClassWithDefaultValues.Long, record.Long.Value);
                    Assert.Equal(testClassWithDefaultValues.Text, record.Text);
                }
            }

            File.Delete(filepath);
        }

        [Fact]
        public void TestWrite()
        {
            var filepath = base.ConstructTempExcelFilePath();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapDefaults>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(base.SpreadsheetValidator.HasErrors);

            using (var spreadsheet = new Spreadsheet(filepath))
            {
                foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMap>("Sheet1"))
                {
                    Assert.Equal(testClassWithDefaultValues.Bool, record.Bool.Value);
                    Assert.Equal(testClassWithDefaultValues.Byte, record.Byte.Value);
                    Assert.Equal(testClassWithDefaultValues.Char, record.Char.Value);
                    Assert.Equal(testClassWithDefaultValues.DateTime, record.DateTime.Value);
                    Assert.Equal(testClassWithDefaultValues.Decimal, record.Decimal.Value);
                    Assert.Equal(testClassWithDefaultValues.Double, record.Double.Value);
                    Assert.Equal(testClassWithDefaultValues.Float, record.Float.Value);
                    Assert.Equal(testClassWithDefaultValues.Int, record.Int.Value);
                    Assert.Equal(testClassWithDefaultValues.Long, record.Long.Value);
                    Assert.Equal(testClassWithDefaultValues.Text, record.Text);
                }
            }

            File.Delete(filepath);
        }

        public class TestClass
        {
            public bool? Bool { get; set; }
            public byte? Byte { get; set; }
            public char? Char { get; set; }
            public DateTime? DateTime { get; set; }
            public decimal? Decimal { get; set; }
            public double? Double { get; set; }
            public float? Float { get; set; }
            public int? Int { get; set; }
            public long? Long { get; set; }
            public string Text { get; set; }
        }

        private class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.Bool).Index(1);
                base.Map(x => x.Byte).Index(2);
                base.Map(x => x.Char).Index(3);
                base.Map(x => x.DateTime).Index(4);
                base.Map(x => x.Decimal).Index(5);
                base.Map(x => x.Double).Index(6);
                base.Map(x => x.Float).Index(7);
                base.Map(x => x.Int).Index(8);
                base.Map(x => x.Long).Index(9);
                base.Map(x => x.Text).Index(10);
            }
        }

        private class TestClassMapDefaults : ClassMap<TestClass>
        {
            public TestClassMapDefaults()
            {
                base.Map(x => x.Bool).Index(1).Default(testClassWithDefaultValues.Bool);
                base.Map(x => x.Byte).Index(2).Default(testClassWithDefaultValues.Byte);
                base.Map(x => x.Char).Index(3).Default(testClassWithDefaultValues.Char);
                base.Map(x => x.DateTime).Index(4).Default(testClassWithDefaultValues.DateTime);
                base.Map(x => x.Decimal).Index(5).Default(testClassWithDefaultValues.Decimal);
                base.Map(x => x.Double).Index(6).Default(testClassWithDefaultValues.Double);
                base.Map(x => x.Float).Index(7).Default(testClassWithDefaultValues.Float);
                base.Map(x => x.Int).Index(8).Default(testClassWithDefaultValues.Int);
                base.Map(x => x.Long).Index(9).Default(testClassWithDefaultValues.Long);
                base.Map(x => x.Text).Index(10).Default(testClassWithDefaultValues.Text);
            }
        }
    }
}