namespace Tests.Styles
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    using Xunit;

    using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

    public class WorksheetStyles : SpreadsheetTesterBase
    {
        private static readonly Dictionary<uint, string> cellValues = new Dictionary<uint, string>()
        {
            { 1, "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog." },
            { 2, "B" },
        };

        private static readonly Dictionary<uint, string> headerNames = new Dictionary<uint, string>()
        {
            { 1, "Header1" },
            { 2, "Header2" },
        };

		[Fact]
		public void TestAutoFilter()
		{
			var filepath = base.ConstructTempXlsxSaveName();
			using (var spreadsheet = new Spreadsheet(filepath))
			{
				spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { HeaderRowIndex = 1, ShouldAutoFilter = true });
				spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", base.CreateRecords<TestClass>(10), new WorksheetStyle() { HeaderRowIndex = 2, ShouldAutoFilter = true });
				spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet3", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldAutoFilter = false });
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

					var worksheet1 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet1".Equals(s.Name)).Id)).Worksheet;
					var autoFilter1 = worksheet1.Elements<OpenXml.AutoFilter>().FirstOrDefault();
					Assert.Equal("A1:B1", autoFilter1.Reference);

					var worksheet2 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet2".Equals(s.Name)).Id)).Worksheet;
					var autoFilter2 = worksheet2.Elements<OpenXml.AutoFilter>().FirstOrDefault();
					Assert.Equal("A2:B2", autoFilter2.Reference);

					var worksheet3 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet3".Equals(s.Name)).Id)).Worksheet;
					var autoFilter3 = worksheet3.Elements<OpenXml.AutoFilter>().FirstOrDefault();
					Assert.Null(autoFilter3);
				}

				File.Delete(spreadsheetFile);
			}
		}

		[Fact]
        public void TestColumnWidths()
        {
			const double maxWidth = 20;
			const double minWidth = 10;

            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { MaxColumnWidth = maxWidth, MinColumnWidth = minWidth, ShouldAutoFitColumns = true });
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
					var worksheetPart = (WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet1".Equals(s.Name)).Id);
					var columns = worksheetPart.Worksheet.GetFirstChild<OpenXml.Columns>();
					foreach (OpenXml.Column column in columns.ChildElements)
					{
						Assert.True(column.Width >= minWidth && column.Width <= maxWidth);
					}
				}

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestShouldFreezeHeaderRow()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { HeaderRowIndex = 1, ShouldFreezeHeaderRow = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", base.CreateRecords<TestClass>(10), new WorksheetStyle() { HeaderRowIndex = 4, ShouldFreezeHeaderRow = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet3", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldFreezeHeaderRow = false });
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

                    var sheet1 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet1".Equals(s.Name)).Id)).Worksheet;
                    var pane1 = sheet1.SheetViews.FirstOrDefault()?.Descendants<OpenXml.Pane>().FirstOrDefault();
                    Assert.Equal<OpenXml.PaneStateValues>(OpenXml.PaneStateValues.Frozen, pane1.State);
                    Assert.Equal<OpenXml.PaneValues>(OpenXml.PaneValues.BottomLeft, pane1.ActivePane);
                    Assert.Equal(1, pane1.VerticalSplit);
                    Assert.Equal("A2", pane1.TopLeftCell);

					var sheet2 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet2".Equals(s.Name)).Id)).Worksheet;
					var pane2 = sheet2.SheetViews.FirstOrDefault()?.Descendants<OpenXml.Pane>().FirstOrDefault();
					Assert.Equal<OpenXml.PaneStateValues>(OpenXml.PaneStateValues.Frozen, pane2.State);
					Assert.Equal<OpenXml.PaneValues>(OpenXml.PaneValues.BottomLeft, pane2.ActivePane);
                    Assert.Equal(4, pane2.VerticalSplit);
                    Assert.Equal("A5", pane2.TopLeftCell);

					var sheet3 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet3".Equals(s.Name)).Id)).Worksheet;
                    var pane3 = sheet3.SheetViews.FirstOrDefault()?.Descendants<OpenXml.Pane>().FirstOrDefault();
                    Assert.Null(pane3);
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestShouldWriteHeader()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false });
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
                    var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    var sheet1 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet1".Equals(s.Name)).Id)).Worksheet;
                    foreach (var kvp in headerNames)
                    {
                        var cell = sheet1.Descendants<OpenXml.Cell>().ElementAt((int)kvp.Key - 1);
                        var sharedStringValue = base.GetSharedStringValue(sharedStringTablePart, cell.CellValue.InnerText);
                        Assert.Equal(kvp.Value, sharedStringValue);
                    }

                    var sheet2 = ((WorksheetPart)workbookPart.GetPartById(workbookPart.Workbook.Descendants<OpenXml.Sheet>().First(s => "Sheet2".Equals(s.Name)).Id)).Worksheet;
                    foreach (var kvp in cellValues)
                    {
                        var cell = sheet2.Descendants<OpenXml.Cell>().ElementAt((int)kvp.Key - 1);
                        var sharedStringValue = base.GetSharedStringValue(sharedStringTablePart, cell.CellValue.InnerText);
                        Assert.Equal(kvp.Value, sharedStringValue);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        private class TestClass
        {
            public string LongText { get; set; } = cellValues[1];
            public string ShortText { get; set; } = cellValues[2];
        }

        private class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.LongText).Index(1).Name(headerNames[1]);
                base.Map(x => x.ShortText).Index(2).Name(headerNames[2]);
            }
        }
    }
}