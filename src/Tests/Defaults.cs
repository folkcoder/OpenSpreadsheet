namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    using Tests.Models;

    using Xunit;

    public class Defaults
    {
        private const bool defaultBool = false;
        private const byte defaultByte = 4;
        private const char defaultChar = 'a';
        private const decimal defaultDecimal = 999.99M;
        private const double defaultDouble = 2.8E+8;
        private const float defaultFloat = 99.9F;
        private const int defaultInt = -1212454;
        private const long defaultLong = 999999999999999999;
        private const string defaultText = "default string";
        private const long defaultTicks = 636826572370000000;

        private const string defaultsReadSheetName = "read";
        private const string defaultsWriteSheetName = "write";

        private const int recordCount = 25;
        private readonly string filepath;

        public Defaults()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "defaults.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(recordCount);
                spreadsheet.WriteWorksheet<NullableDataTypes, TestClassMapEmpty>(defaultsReadSheetName, records);
                spreadsheet.WriteWorksheet<NullableDataTypes, TestClassMapDefaults>(defaultsWriteSheetName, records);
            }
        }

        [Fact]
        public void TestRead()
        {
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = spreadsheet.ReadWorksheet<NullableDataTypes, TestClassMapEmpty>(defaultsReadSheetName).ToList();
                Assert.Equal(records.Count, recordCount);
                foreach (var record in records)
                {
                    Assert.Equal(record.Bool, defaultBool);
                    Assert.Equal(record.Byte, defaultByte);
                    Assert.Equal(record.Char, defaultChar);
                    Assert.Equal(record.DateTime, new DateTime(defaultTicks));
                    Assert.Equal(record.Decimal, defaultDecimal);
                    Assert.Equal(record.Double, defaultDouble);
                    Assert.Equal(record.Float, defaultFloat);
                    Assert.Equal(record.Int, defaultInt);
                    Assert.Equal(record.Long, defaultLong);
                    Assert.Equal(record.Text, defaultText);
                }
            }
        }

        [Fact]
        public void TestWrite()
        {
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = spreadsheet.ReadWorksheet<NullableDataTypes, TestClassMapDefaults>(defaultsWriteSheetName).ToList();
                Assert.Equal(records.Count, recordCount);
                foreach (var record in records)
                {
                    Assert.Equal(record.Bool, defaultBool);
                    Assert.Equal(record.Byte, defaultByte);
                    Assert.Equal(record.Char, defaultChar);
                    Assert.Equal(record.DateTime, new DateTime(defaultTicks));
                    Assert.Equal(record.Decimal, defaultDecimal);
                    Assert.Equal(record.Double, defaultDouble);
                    Assert.Equal(record.Float, defaultFloat);
                    Assert.Equal(record.Int, defaultInt);
                    Assert.Equal(record.Long, defaultLong);
                    Assert.Equal(record.Text, defaultText);
                }
            }
        }

        [Fact]
        public void TestSpreadsheetWriting()
        {
            var validator = new SpreadsheetValidator();
            validator.Validate(this.filepath);

            Assert.False(validator.HasErrors);
        }

        private static IEnumerable<NullableDataTypes> CreateTestRecords(int count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return new NullableDataTypes();
            }
        }

        internal class TestClassMapEmpty : ClassMap<NullableDataTypes>
        {
            public TestClassMapEmpty()
            {
                base.Map(x => x.Bool).Index(1).ConstantWrite(string.Empty).DefaultRead(defaultBool);
                base.Map(x => x.Byte).Index(2).ConstantWrite(string.Empty).DefaultRead(defaultByte);
                base.Map(x => x.Char).Index(3).ConstantWrite(string.Empty).DefaultRead(defaultChar);
                base.Map(x => x.DateTime).Index(4).ConstantWrite(string.Empty).DefaultRead(new DateTime(defaultTicks));
                base.Map(x => x.Decimal).Index(5).ConstantWrite(string.Empty).DefaultRead(defaultDecimal);
                base.Map(x => x.Double).Index(6).ConstantWrite(string.Empty).DefaultRead(defaultDouble);
                base.Map(x => x.Float).Index(7).ConstantWrite(string.Empty).DefaultRead(defaultFloat);
                base.Map(x => x.Int).Index(8).ConstantWrite(string.Empty).DefaultRead(defaultInt);
                base.Map(x => x.Long).Index(9).ConstantWrite(string.Empty).DefaultRead(defaultLong);
                base.Map(x => x.Text).Index(10).ConstantWrite(string.Empty).DefaultRead(defaultText);
            }
        }

        internal class TestClassMapDefaults : ClassMap<NullableDataTypes>
        {
            public TestClassMapDefaults()
            {
                base.Map(x => x.Bool).Index(1).Default(defaultBool);
                base.Map(x => x.Byte).Index(2).Default(defaultByte);
                base.Map(x => x.Char).Index(3).Default(defaultChar);
                base.Map(x => x.DateTime).Index(4).Default(new DateTime(defaultTicks));
                base.Map(x => x.Decimal).Index(5).Default(defaultDecimal);
                base.Map(x => x.Double).Index(6).Default(defaultDouble);
                base.Map(x => x.Float).Index(7).Default(defaultFloat);
                base.Map(x => x.Int).Index(8).Default(defaultInt);
                base.Map(x => x.Long).Index(9).Default(defaultLong);
                base.Map(x => x.Text).Index(10).Default(defaultText);
            }
        }
    }
}