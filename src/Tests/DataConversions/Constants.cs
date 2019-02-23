namespace Tests.DataConversions
{
    using System;
    using System.IO;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    using Xunit;

    public class Constants : SpreadsheetTesterBase
    {
        private static readonly TestClass testClassWithConstantValues = new TestClass()
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
            Text = "constant string",
        };

        [Fact]
        public void TestRead()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapWithoutConstants>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var spreadsheet = new Spreadsheet(spreadsheetFile))
                {
                    foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMapWithConstants>("Sheet1"))
                    {
                        Assert.Equal(record.Bool, testClassWithConstantValues.Bool);
                        Assert.Equal(record.Byte, testClassWithConstantValues.Byte);
                        Assert.Equal(record.Char, testClassWithConstantValues.Char);
                        Assert.Equal(record.DateTime, testClassWithConstantValues.DateTime);
                        Assert.Equal(record.Decimal, testClassWithConstantValues.Decimal);
                        Assert.Equal(record.Double, testClassWithConstantValues.Double);
                        Assert.Equal(record.Float, testClassWithConstantValues.Float);
                        Assert.Equal(record.Int, testClassWithConstantValues.Int);
                        Assert.Equal(record.Long, testClassWithConstantValues.Long);
                        Assert.Equal(record.Text, testClassWithConstantValues.Text);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestWrite()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapWithConstants>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var spreadsheet = new Spreadsheet(spreadsheetFile))
                {
                    foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMapWithoutConstants>("Sheet1"))
                    {
                        Assert.Equal(record.Bool, testClassWithConstantValues.Bool);
                        Assert.Equal(record.Byte, testClassWithConstantValues.Byte);
                        Assert.Equal(record.Char, testClassWithConstantValues.Char);
                        Assert.Equal(record.DateTime, testClassWithConstantValues.DateTime);
                        Assert.Equal(record.Decimal, testClassWithConstantValues.Decimal);
                        Assert.Equal(record.Double, testClassWithConstantValues.Double);
                        Assert.Equal(record.Float, testClassWithConstantValues.Float);
                        Assert.Equal(record.Int, testClassWithConstantValues.Int);
                        Assert.Equal(record.Long, testClassWithConstantValues.Long);
                        Assert.Equal(record.Text, testClassWithConstantValues.Text);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        public class TestClass
        {
            public bool Bool { get; set; }
            public byte Byte { get; set; }
            public char Char { get; set; } = 'a';
            public DateTime DateTime { get; set; }
            public decimal Decimal { get; set; }
            public double Double { get; set; }
            public float Float { get; set; }
            public int Int { get; set; }
            public long Long { get; set; }
            public string Text { get; set; } = string.Empty;
        }

        public class TestClassMapWithConstants : ClassMap<TestClass>
        {
            public TestClassMapWithConstants()
            {
                base.Map(x => x.Bool).Index(1).Constant(testClassWithConstantValues.Bool);
                base.Map(x => x.Byte).Index(2).Constant(testClassWithConstantValues.Byte);
                base.Map(x => x.Char).Index(3).Constant(testClassWithConstantValues.Char);
                base.Map(x => x.DateTime).Index(4).Constant(testClassWithConstantValues.DateTime);
                base.Map(x => x.Decimal).Index(5).Constant(testClassWithConstantValues.Decimal);
                base.Map(x => x.Double).Index(6).Constant(testClassWithConstantValues.Double);
                base.Map(x => x.Float).Index(7).Constant(testClassWithConstantValues.Float);
                base.Map(x => x.Int).Index(8).Constant(testClassWithConstantValues.Int);
                base.Map(x => x.Long).Index(9).Constant(testClassWithConstantValues.Long);
                base.Map(x => x.Text).Index(10).Constant(testClassWithConstantValues.Text);
            }
        }

        public class TestClassMapWithoutConstants : ClassMap<TestClass>
        {
            public TestClassMapWithoutConstants()
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
    }
}