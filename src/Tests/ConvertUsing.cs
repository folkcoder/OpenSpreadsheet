namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class ConvertUsing
    {
        private const int recordCount = 25;
        private readonly string filepath;

        public ConvertUsing()
        {
            var folderPath = Path.Combine(Environment.CurrentDirectory, "test_outputs");
            var directory = Directory.CreateDirectory(folderPath);
            this.filepath = Path.Combine(folderPath, "convert_using.xlsx");
            if (File.Exists(this.filepath))
            {
                File.Delete(this.filepath);
            }

            var records = CreateTestRecords(recordCount);
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("test", records);
            }
        }

        [Fact]
        public void TestConvertUsing()
        {
            using (var spreadsheet = new Spreadsheet(this.filepath))
            {
                foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMap>("test"))
                {
                    Assert.Equal(TestEnum.A, record.TestEnum1);
                    Assert.Equal(TestEnum.B, record.TestEnum2);
                    Assert.Equal(TestEnum.C, record.TestEnum3);
                }
            }
        }

        internal class TestClass
        {
            public TestEnum TestEnum1 { get; set; } = TestEnum.A;
            public TestEnum TestEnum2 { get; set; } = TestEnum.B;
            public TestEnum TestEnum3 { get; set; } = TestEnum.C;
        }

        internal enum TestEnum
        {
            Unset = 0,
            A = 1,
            B = 2,
            C = 3
        }

        internal class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                Map(x => x.TestEnum1).Name("No Conversion");

                Map(x => x.TestEnum2).Name("Conversion B")
                    .ReadUsing(row => ConvertStringToEnum(row.GetCellValue("Conversion B")))
                    .WriteUsing(x => ConvertEnumToString(x.TestEnum2));

                Map(x => x.TestEnum3).Name("Conversion C")
                    .ReadUsing(row => ConvertStringToEnum(row.GetCellValue("Conversion C")))
                    .WriteUsing(x => ConvertEnumToString(x.TestEnum3));
            }
        }

        private static IEnumerable<TestClass> CreateTestRecords(int count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return new TestClass();
            }
        }

        private static string ConvertEnumToString(TestEnum testEnum)
        {
            switch (testEnum)
            {
                case TestEnum.A:
                    return "Enum A";

                case TestEnum.B:
                    return "Enum B";

                case TestEnum.C:
                    return "Enum C";

                default:
                    return "Undefined";
            }
        }

        private static TestEnum ConvertStringToEnum(string enumText)
        {
            switch (enumText)
            {
                case "Enum A":
                    return TestEnum.A;

                case "Enum B":
                    return TestEnum.B;

                case "Enum C":
                    return TestEnum.C;

                default:
                    return TestEnum.Unset;
            }
        }
    }
}