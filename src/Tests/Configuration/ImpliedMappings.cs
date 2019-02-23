namespace Tests.Configuration
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class ImpliedMappings : SpreadsheetTesterBase
    {
        [Fact]
        public void TestReadByHeaderName()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            var fileSavedByExcel = base.SaveAsExcelFile(filepath);

            foreach (var spreadsheetFile in new[] { filepath, fileSavedByExcel })
            {
                base.SpreadsheetValidator.Validate(spreadsheetFile);
                Assert.False(base.SpreadsheetValidator.HasErrors);

                using (var spreadsheet = new Spreadsheet(spreadsheetFile))
                {
                    foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMapReadByHeaderName>("Sheet1"))
                    {
                        Assert.Equal(1, record.TestData1);
                        Assert.Equal(2, record.TestData2);
                        Assert.Equal(3, record.TestData3);
                        Assert.Equal(4, record.TestData4);
                        Assert.Equal(5, record.TestData5);
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        [Fact]
        public void TestWriteByMappingOrder()
        {
            var writeOrder = new Dictionary<uint, int>()
            {
                { 1, 1 },
                { 2, 5 },
                { 3, 4 },
                { 4, 3 },
                { 5, 2 },
            };

            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMapWriteByOrder>("Sheet1", base.CreateRecords<TestClass>(10), new WorksheetStyle() { ShouldWriteHeaderRow = false} );
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

                    foreach (var cell in sheet.Descendants<Cell>())
                    {
                        var columnIndex = base.GetColumnIndexFromCellReference(cell.CellReference);
                        var expectedValue = writeOrder[columnIndex];

                        Assert.Equal(expectedValue, int.Parse(cell.InnerText));
                    }
                }

                File.Delete(spreadsheetFile);
            }
        }

        private class TestClass
        {
            public int TestData1 { get; set; } = 1;
            public int TestData2 { get; set; } = 2;
            public int TestData3 { get; set; } = 3;
            public int TestData4 { get; set; } = 4;
            public int TestData5 { get; set; } = 5;
        }

        private class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.TestData1).Index(5);
                base.Map(x => x.TestData2).Index(4);
                base.Map(x => x.TestData3).Index(3);
                base.Map(x => x.TestData4).Index(2);
                base.Map(x => x.TestData5).Index(1);
            }
        }

        private class TestClassMapReadByHeaderName : ClassMap<TestClass>
        {
            public TestClassMapReadByHeaderName()
            {
                base.Map(x => x.TestData5);
                base.Map(x => x.TestData4);
                base.Map(x => x.TestData3);
                base.Map(x => x.TestData2);
                base.Map(x => x.TestData1);
            }
        }

        private class TestClassMapWriteByOrder : ClassMap<TestClass>
        {
            public TestClassMapWriteByOrder()
            {
                base.Map(x => x.TestData5);
                base.Map(x => x.TestData4);
                base.Map(x => x.TestData3);
                base.Map(x => x.TestData2);
                base.Map(x => x.TestData1).Index(1);
            }
        }
    }
}