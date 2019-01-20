namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using SpreadsheetHelper;
    using SpreadsheetHelper.Configuration;

    using Xunit;

    public class ReadWriteOptions
    {
        private const string allRecordsNoWriterSheetName = "all records no explicit writer";
        private const string allRecordsWithWriterSheetName = "all records explicit writer";
        private const string headerRowNotDefault = "custom header row";
        private const string skippedRowsSheetName = "skipped rows";

        private const uint headerRowIndex = 5;
        private const int recordCount = 25;
        private const int rowsToSkip = 5;
        private readonly string filepath;

        public ReadWriteOptions()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "excel tests");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "read_write_options.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var records = CreateTestRecords(recordCount);

                spreadsheet.WriteWorksheet<TestClass, TestClassMap>(allRecordsNoWriterSheetName, records);

                using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>(allRecordsWithWriterSheetName))
                {
                    writer.WriteRecords(records);
                }

                using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>(skippedRowsSheetName))
                {
                    writer.WriteRecord(new TestClass() { TestData = "first row" });
                    writer.WriteRecord(new TestClass() { TestData = "second row" });
                    writer.SkipRows(5);
                    writer.WriteRecord(new TestClass() { TestData = "third row" });
                    writer.WriteRecord(new TestClass() { TestData = "fourth row" });
                }

                using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("manually write header"))
                {
                    writer.WriteHeader();
                    writer.WriteRecords(records);
                    writer.WriteHeader();
                }

                spreadsheet.WriteWorksheet<TestClass, TestClassMap>(headerRowNotDefault, records, new WorksheetStyle() { HeaderRowIndex = headerRowIndex });

                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("specified worksheet index", 0, records);
            }
        }

        [Fact]
        public void TestRead()
        {
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var recordsNoExplicitReader = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(allRecordsNoWriterSheetName);
                Assert.Equal(recordCount, recordsNoExplicitReader.Count());

                using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>(allRecordsWithWriterSheetName))
                {
                    var recordsExplicitReader = reader.ReadRows();
                    Assert.Equal(recordCount, recordsExplicitReader.Count());
                }

                using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>(skippedRowsSheetName))
                {
                    var firstRow = reader.ReadRow();
                    Assert.Equal("first row", firstRow.TestData);

                    var secondRow = reader.ReadRow();
                    Assert.Equal("second row", secondRow.TestData);

                    reader.SkipRows(rowsToSkip);

                    var thirdRow = reader.ReadRow();
                    Assert.Equal("third row", thirdRow.TestData);

                    var fourthRow = reader.ReadRow();
                    Assert.Equal("fourth row", fourthRow.TestData);
                }

                var recordsCustomHeaderRow = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(headerRowNotDefault, headerRowIndex);
                Assert.Equal(recordCount, recordsCustomHeaderRow.Count());
            }
        }

        [Fact]
        public void TestWrite()
        {
            var validator = new SpreadsheetValidator();
            validator.Validate(this.filepath);

            Assert.False(validator.HasErrors);
        }

        [Fact]
        public void AddSheetToExistingWorkbook()
        {
            const string newWorksheetName = "new worksheet";

            var records = CreateTestRecords(25);
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>(newWorksheetName, records);
            }

            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                var readRecords = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(newWorksheetName);
                Assert.Equal(recordCount, readRecords.Count());
            }
        }

        [Fact]
        public void AddWorksheetsWithSameName()
        {
            const string newWorksheetName = "worksheet name";

            var records = CreateTestRecords(25);
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>(newWorksheetName, records);

                try
                {
                    spreadsheet.WriteWorksheet<TestClass, TestClassMap>(newWorksheetName, records);
                }
                catch (ArgumentException)
                {
                    Assert.True(true);
                    return;
                }
            }

            Assert.True(false);
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

        internal class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                Map(x => x.TestData).Index(1);
            }
        }
    }
}