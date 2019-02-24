namespace Tests.SpreadsheetIO
{
	using System;
	using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class Reader : SpreadsheetTesterBase
	{
        [Fact]
        public void TestReadAll()
        {
            const int recordCount = 100;

            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(recordCount));
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                using (var spreadsheet = new Spreadsheet(spreadsheetFile))
                {
                    var records = spreadsheet.ReadWorksheet<TestClass, TestClassMap>("Sheet1").ToList();
                    Assert.Equal(recordCount, records.Count);

                    foreach (var record in records)
                    {
                        Assert.Equal("test data", record.TestData);
                    }

                    using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>("Sheet1"))
                    {
                        var records2 = spreadsheet.ReadWorksheet<TestClass, TestClassMap>("Sheet1").ToList();
                        Assert.Equal(recordCount, records2.Count);

                        foreach (var record in records2)
                        {
                            Assert.Equal("test data", record.TestData);
                        }
                    }
                }

                File.Delete(spreadsheetFile);
            }

            //using (var spreadsheet = new Spreadsheet(this.filepath))
            //{
            //    var recordsNoExplicitReader = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(allRecordsNoWriterSheetName);
            //    Assert.Equal(recordCount, recordsNoExplicitReader.Count());

            //    using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>(allRecordsWithWriterSheetName))
            //    {
            //        var recordsExplicitReader = reader.ReadRows();
            //        Assert.Equal(recordCount, recordsExplicitReader.Count());
            //    }

            //    using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>(skippedRowsSheetName))
            //    {
            //        var firstRow = reader.ReadRow();
            //        Assert.Equal("first row", firstRow.TestData);

            //        var secondRow = reader.ReadRow();
            //        Assert.Equal("second row", secondRow.TestData);

            //        reader.SkipRows(rowsToSkip);

            //        var thirdRow = reader.ReadRow();
            //        Assert.Equal("third row", thirdRow.TestData);

            //        var fourthRow = reader.ReadRow();
            //        Assert.Equal("fourth row", fourthRow.TestData);
            //    }

            //    var recordsCustomHeaderRow = spreadsheet.ReadWorksheet<TestClass, TestClassMap>(headerRowNotDefault, headerRowIndex);
            //    Assert.Equal(recordCount, recordsCustomHeaderRow.Count());
            //}
        }

        private class TestClass
		{
			public string TestData { get; set; } = "test data";
		}

		private class TestClassMap : ClassMap<TestClass>
		{
			public TestClassMap() => base.Map(x => x.TestData).Index(1);
		}
	}
}