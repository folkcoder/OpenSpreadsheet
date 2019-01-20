namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using SpreadsheetHelper;
    using SpreadsheetHelper.Configuration;

    using Tests.Models;

    using Xunit;

    public class Constants
    {
        private const bool constantBool = true;
        private const byte constantByte = 24;
        private const char constantChar = 'c';
        private const decimal constantDecimal = 433123.12M;
        private const double constantDouble = 6.1E+3;
        private const float constantFloat = 23342.93F;
        private const int constantInt = 53242;
        private const long constantLong = 4324823423423;
        private const string constantsSheetName = "Constants";
        private const string constantText = "constant string";
        private const long constantTicks = 636826537190000000;
        private const int recordCount = 25;
        private readonly string filepath;

        public Constants()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "excel tests");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "constants.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(filepath))
            {
                var records = CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<DataTypes, TestClassMapConstants>(constantsSheetName, records);
            }
        }

        [Fact]
        public void TestRead()
        {
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                var records = spreadsheet.ReadWorksheet<DataTypes, TestClassMapConstants>(constantsSheetName).ToList();
                Assert.Equal(records.Count, recordCount);
                foreach (var record in records)
                {
                    Assert.Equal(record.Bool, constantBool);
                    Assert.Equal(record.Byte, constantByte);
                    Assert.Equal(record.Char, constantChar);
                    Assert.Equal(record.DateTime, new DateTime(constantTicks));
                    Assert.Equal(record.Decimal, constantDecimal);
                    Assert.Equal(record.Double, constantDouble);
                    Assert.Equal(record.Float, constantFloat);
                    Assert.Equal(record.Int, constantInt);
                    Assert.Equal(record.Long, constantLong);
                    Assert.Equal(record.Text, constantText);
                }
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

        internal class TestClassMapConstants : ClassMap<DataTypes>
        {
            public TestClassMapConstants()
            {
                Map(x => x.Bool).Index(1).Constant(constantBool);
                Map(x => x.Byte).Index(2).Constant(constantByte);
                Map(x => x.Char).Index(3).Constant(constantChar);
                Map(x => x.DateTime).Index(4).Constant(new DateTime(constantTicks));
                Map(x => x.Decimal).Index(5).Constant(constantDecimal);
                Map(x => x.Double).Index(6).Constant(constantDouble);
                Map(x => x.Float).Index(7).Constant(constantFloat);
                Map(x => x.Int).Index(8).Constant(constantInt);
                Map(x => x.Long).Index(9).Constant(constantLong);
                Map(x => x.Text).Index(10).Constant(constantText);
            }
        }
    }
}