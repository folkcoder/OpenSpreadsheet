namespace Benchmarks.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using BenchmarkDotNet.Attributes;
    using ClosedXML.Excel;
    using OfficeOpenXml;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;

    [MemoryDiagnoser]
    public class BenchmarkRead
    {
        private const string worksheetName = "test sheet";
        private string inputPath;

        [Params(50000, 100000, 250000, 500000)]
        public int RecordCount { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            this.inputPath = Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx";
            var records = CreateRecords(this.RecordCount);
            using (var spreadsheet = new Spreadsheet(this.inputPath))
            {
                spreadsheet.WriteWorksheet<TestRecord, TestClassMap>(worksheetName, records);
            }
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            File.Delete(this.inputPath);
        }

        [Benchmark]
        public void TestClosedXml()
        {
            var records = new List<TestRecord>();
            using (var workbook = new XLWorkbook(this.inputPath, XLEventTracking.Disabled))
            {
                workbook.TryGetWorksheet(worksheetName, out IXLWorksheet worksheet);

                for (int i = 0; i < this.RecordCount; i++)
                {
                    int row = i + 1;
                    records.Add(new TestRecord
                    {
                        TestData1 = (string)worksheet.Cell(row, 1).Value,
                        TestData2 = (string)worksheet.Cell(row, 2).Value,
                        TestData3 = (string)worksheet.Cell(row, 3).Value
                    });
                }
            }
        }

        [Benchmark]
        public void TestEPPlus()
        {
            var records = new List<TestRecord>();
            var fileInfo = new FileInfo(this.inputPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                var worksheet = excelPackage.Workbook.Worksheets[worksheetName];

                for (int i = 0; i < this.RecordCount; i++)
                {
                    int row = i + 1;
                    records.Add(new TestRecord
                    {
                        TestData1 = (string)worksheet.Cells[row, 1].Value,
                        TestData2 = (string)worksheet.Cells[row, 2].Value,
                        TestData3 = (string)worksheet.Cells[row, 3].Value
                    });
                }
            }
        }

        [Benchmark]
        public void TestOpenSpreadsheet()
        {
            using (var spreadsheet = new Spreadsheet(this.inputPath))
            {
                var records = spreadsheet.ReadWorksheet<TestRecord, TestClassMap>(worksheetName).ToList();
            }
        }

        internal class TestClassMap : ClassMap<TestRecord>
        {
            public TestClassMap()
            {
                Map(x => x.TestData1).Index(1);
                Map(x => x.TestData2).Index(2);
                Map(x => x.TestData3).Index(3);
            }
        }

        internal class TestRecord
        {
            public string TestData1 { get; set; }
            public string TestData2 { get; set; }
            public string TestData3 { get; set; }
        }

        private static IEnumerable<TestRecord> CreateRecords(int amount)
        {
            for (int i = 0; i < amount; i++)
            {
                yield return new TestRecord()
                {
                    TestData1 = "sadassdgsdfsdfsgsdfsdfds",
                    TestData2 = "afasdljfsdlgjsdljfsldjfl",
                    TestData3 = "g;fdgjlsdfhsdfiefndslc k",
                };
            }
        }
    }
}