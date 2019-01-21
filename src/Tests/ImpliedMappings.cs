namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using SpreadsheetHelper;
    using SpreadsheetHelper.Configuration;
    using Xunit;

    public class ImpliedMappings
    {
        private const string unspecifiedIndexesSheetName = "unspecified indexes";

        private const int recordCount = 25;
        private readonly string filepath;

        public ImpliedMappings()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "implied_mappings.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            var records = CreateTestRecords(recordCount);
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapUnspecifiedIndexes>(unspecifiedIndexesSheetName, records);
            }
        }

        //[Fact]
        //public void TestUnspecifiedIndexesValidation()
        //{
        //    using (var spreadsheet = new Spreadsheet(this.filepath))
        //    {
        //        using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMapUnspecifiedIndexes>(unspecifiedIndexesSheetName))
        //        {
        //        }

        //        var recordsCustomHeaderRow = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(headerRowNotDefault, headerRowIndex);
        //        Assert.Equal(recordCount, recordsCustomHeaderRow.Count());
        //    }
        //}

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
            public string TestData1 { get; set; }
            public string TestData2 { get; set; }
            public string TestData3 { get; set; }
            public string TestData4 { get; set; }
            public string TestData5 { get; set; }
        }

        internal class TestClassMapUnspecifiedIndexes : ClassMap<TestClass>
        {
            public TestClassMapUnspecifiedIndexes()
            {
                Map(x => x.TestData5).IndexWrite(5).ConstantWrite("5");
                Map(x => x.TestData2).ConstantWrite("2");
                Map(x => x.TestData1).IndexWrite(1).ConstantWrite("1");
                Map(x => x.TestData3).ConstantWrite("3");
                Map(x => x.TestData4).ConstantWrite("4");
            }
        }
    }
}
