namespace Tests.DataConversions
{
    using System.IO;
    using OpenSpreadsheet;
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class ConvertUsing : SpreadsheetTesterBase
    {
        [Fact]
        public void TestConvertUsing()
        {
            var filepath = base.ConstructTempXlsxSaveName();
            using (var spreadsheet = new Spreadsheet(filepath))
            {
                spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", base.CreateRecords<TestClass>(10));
            }

            base.SpreadsheetValidator.Validate(filepath);
            Assert.False(base.SpreadsheetValidator.HasErrors);

            using (var spreadsheet = new Spreadsheet(filepath))
            {
                foreach (var record in spreadsheet.ReadWorksheet<TestClass, TestClassMap>("Sheet1"))
                {
                    Assert.Equal(TestEnum.A, record.TestEnum1);
                    Assert.Equal(TestEnum.B, record.TestEnum2);
                    Assert.Equal(TestEnum.C, record.TestEnum3);
                }
            }

            File.Delete(filepath);
        }

        private class TestClass
        {
            public TestEnum TestEnum1 { get; set; } = TestEnum.A;
            public TestEnum TestEnum2 { get; set; } = TestEnum.B;
            public TestEnum TestEnum3 { get; set; } = TestEnum.C;
        }

        private enum TestEnum
        {
            Unset = 0,
            A,
            B,
            C
        }

        private class TestClassMap : ClassMap<TestClass>
        {
            public TestClassMap()
            {
                base.Map(x => x.TestEnum1).Name("No Conversion");

                base.Map(x => x.TestEnum2).Name("Conversion B")
                    .ReadUsing(row => ConvertStringToEnum(row.GetCellValue("Conversion B")))
                    .WriteUsing(x => ConvertEnumToString(x.TestEnum2));

                base.Map(x => x.TestEnum3).Name("Conversion C")
                    .ReadUsing(row => ConvertStringToEnum(row.GetCellValue("Conversion C")))
                    .WriteUsing(x => ConvertEnumToString(x.TestEnum3));
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