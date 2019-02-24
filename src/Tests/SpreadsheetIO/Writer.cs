namespace Tests.SpreadsheetIO
{
    using System;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class Writer : SpreadsheetTesterBase
	{
		[Fact]
		public void AddWorksheetsWithSameName()
		{
			var records = base.CreateRecords<TestClass>(10);
			var filepath = base.ConstructTempXlsxSaveName();
			using (var spreadsheet = new Spreadsheet(filepath))
			{
				spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10));
				Assert.Throws<ArgumentException>(() => spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", records));
			}
		}

		[Fact]
		public void WriteAllRecords()
		{
			const int recordCount = 100;

			var records = base.CreateRecords<TestClass>(recordCount);
			var filepath = base.ConstructTempXlsxSaveName();
			using (var spreadsheet = new Spreadsheet(filepath))
			{
				spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", records, new WorksheetStyle() { ShouldWriteHeaderRow = false } );

				using (var explicitWriter = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("Sheet2", new WorksheetStyle() { ShouldWriteHeaderRow = false }))
				{
					foreach (var record in records)
					{
						explicitWriter.WriteRecord(record);
					}
				}
			}

			var fileSavedByExcel = base.SaveAsExcelFile(filepath);
			foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
			{
				base.SpreadsheetValidator.Validate(spreadsheetFile);
				Assert.False(base.SpreadsheetValidator.HasErrors);

				using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
				using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
				{
					var workbookPart = spreadsheetDocument.WorkbookPart;
					foreach (var worksheetPart in workbookPart.WorksheetParts)
					{
						var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
						Assert.Equal(recordCount, sheetData.Elements<Row>().Count());
					}
				}

				File.Delete(spreadsheetFile);
			}
		}

        [Fact]
        public void TestSkipRows()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("Sheet1", new WorksheetStyle() { ShouldWriteHeaderRow = false }))
            {
                writer.WriteRecord(new TestClass());
                writer.SkipRows(5);
                writer.WriteRecord(new TestClass());
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    var rows = sheet.Descendants<Row>();
                    Assert.Equal<uint>(1, rows.ElementAt(0).RowIndex.Value);
                    Assert.Equal<uint>(7, rows.ElementAt(1).RowIndex.Value);
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestWorksheetIndex()
        {
            var records = base.CreateRecords<TestClass>(10);
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet3", records);
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", 0, records);
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    Assert.Equal("Sheet1", spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(0).Name);
                    Assert.Equal("Sheet2", spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(1).Name);
                    Assert.Equal("Sheet3", spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(2).Name);
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
		public void TestWriteHeaders()
		{
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("Sheet1", new WorksheetStyle() { ShouldWriteHeaderRow = false }))
            {
                writer.WriteHeader();
                writer.WriteHeader();
                writer.WriteHeader();
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);
            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var filestream = new FileStream(spreadsheetFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var spreadsheetDocument = SpreadsheetDocument.Open(filestream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;
                    var sharedStringPart = workbookPart.SharedStringTablePart;

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var cellValue = base.GetSharedStringValue(sharedStringPart, cell.CellValue.InnerText);
                        Assert.Equal("TestData", cellValue);
                    }
                }

                File.Delete(spreadsheetFile);
            }
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