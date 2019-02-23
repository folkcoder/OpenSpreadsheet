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
            { 1, "The quick brown fox jumps over the lazy dog." },
            { 2, "B" },
        };

        private static readonly Dictionary<uint, string> headerNames = new Dictionary<uint, string>()
        {
            { 1, "Header1" },
            { 2, "Header2" },
        };

        [Fact]
        public void TestShouldFreezeTopRow()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldFreezeHeaderRow = true });
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldFreezeHeaderRow = false });
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

//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Default Style", records);
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Chartreuse Header Background", records, new WorksheetStyle() { HeaderBackgroundColor = Color.Chartreuse, HeaderBackgroundPatternType = OpenXml.PatternValues.Solid });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Comic Sans 16 Italic Font", records, new WorksheetStyle() { HeaderFont = new Font("Comic Sans", 16, FontStyle.Italic) });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Dark Green Header Foreground", records, new WorksheetStyle() { HeaderForegroundColor = Color.DarkGreen });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Centered Horizontal Headers", records, new WorksheetStyle() { HeaderHoizontalAlignment = OpenXml.HorizontalAlignmentValues.Center });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Top Vertical Headers", records, new WorksheetStyle() { HeaderVerticalAlignment = OpenXml.VerticalAlignmentValues.Top });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Min Col Width 10", records, new WorksheetStyle() { MinColumnWidth = 10 });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Max Col Width 5", records, new WorksheetStyle() { MaxColumnWidth = 5 });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Autofilter On", records, new WorksheetStyle() { ShouldAutoFilter = true });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Freeze Top Row", records, new WorksheetStyle() { ShouldFreezeTopRow = true });
//spreadsheet.WriteWorksheet<TestClass, TestClassMap>("AutoFit Columns", records, new WorksheetStyle() { ShouldAutoFitColumns = true });

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