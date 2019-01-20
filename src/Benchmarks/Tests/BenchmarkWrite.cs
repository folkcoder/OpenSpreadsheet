namespace Benchmarks.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using BenchmarkDotNet.Attributes;
    using ClosedXML.Excel;
    using OfficeOpenXml;
    using SpreadsheetHelper;
    using SpreadsheetHelper.Configuration;

    [MemoryDiagnoser]
    public class BenchmarkWrite
    {
        private const uint columnCount = 30;
        private IList<TestRecord> records;

        [Params(50000, 100000, 250000)]
        public int RecordCount { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            this.records = CreateRecords(this.RecordCount).ToList();
        }

        [Benchmark]
        public void TestClosedXml()
        {
            string outputPath = Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx";
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                var worksheet = workbook.Worksheets.Add("Test Sheet");

                int row = 1;
                foreach (var record in this.records)
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        int col = i + 1;
                        worksheet.Cell(row, col).Value = record.TestData;
                    }

                    row++;
                }

                workbook.SaveAs(outputPath);
            }

            File.Delete(outputPath);
        }

        [Benchmark]
        public void TestEPPlus()
        {
            string outputPath = Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx";
            var fileInfo = new FileInfo(outputPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("test sheet");

                int row = 1;
                foreach (var record in this.records)
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        int col = i + 1;
                        worksheet.Cells[row, col].Value = record.TestData;
                    }

                    row++;
                }

                excelPackage.Save();
            }

            File.Delete(outputPath);
        }

        [Benchmark]
        public void TestSpreadsheetHelper()
        {
            string outputPath = Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx";
            using (var spreadsheet = new Spreadsheet(outputPath))
            {
                spreadsheet.WriteWorksheet<TestRecord, TestClassMap>("test sheet", records);
            }

            File.Delete(outputPath);
        }

        internal class TestClassMap : ClassMap<TestRecord>
        {
            public TestClassMap()
            {
                for (uint i = 0; i < columnCount; i++)
                {
                    uint colIndex = i + 1;
                    Map(x => x.TestData).Index(colIndex).IgnoreRead(true);
                }
            }
        }

        internal class TestRecord
        {
            public string TestData { get; set; } = "sadassdgsdfsdfsgsdfsdfds";
        }

        private static IEnumerable<TestRecord> CreateRecords(int amount)
        {
            for (int i = 0; i < amount; i++)
            {
                yield return new TestRecord();
            }
        }
    }
}