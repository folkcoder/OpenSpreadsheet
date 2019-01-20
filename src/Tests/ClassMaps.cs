namespace Tests
{
    using SpreadsheetHelper.Configuration;
    using Xunit;

    public class ClassMaps
    {
        [Fact]
        public void TestConstantsValidation()
        {
            var classMap = new TestClassMapConstants();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
        }

        [Fact]
        public void TestDefaultsValidation()
        {
            var classMap = new TestClassMapDefaults();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
        }

        [Fact]
        public void TestDuplicateIndexesValidation()
        {
            var classMap = new TestClassMapDuplicateIndexes();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.HasErrors);
        }

        [Fact]
        public void TestDuplicateReadPropertiesValidation()
        {
            var classMap = new TestClassMapDuplicateReadProperties();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.HasErrors);
        }

        [Fact]
        public void TestIndexOutOfRangeValidation()
        {
            var classMap = new TestClassMapIndexOutOfRange();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.Errors.Count == 2);
        }

        [Fact]
        public void TestLongHeadersValidation()
        {
            var classMap = new TestClassMapLongHeaders();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
        }

        [Fact]
        public void TestMissingIndexesValidation()
        {
            var classMap = new TestClassMapMissingIndexes();
            var validator = new ConfigurationValidator(classMap.PropertyMaps);
            validator.Validate();
            Assert.True(validator.Errors.Count == 2); // read and write
        }

        internal class TestClass
        {
            public string TestData { get; set; } = "test data";
            public string TestDataNull { get; set; } = null;
        }

        internal class TestClassMapConstants : ClassMap<TestClass>
        {
            public TestClassMapConstants()
            {
                Map(x => x.TestData).Index(1).Constant(2312.231M);
                Map(x => x.TestDataNull).Index(2);
            }
        }

        internal class TestClassMapDefaults : ClassMap<TestClass>
        {
            public TestClassMapDefaults()
            {
                Map(x => x.TestDataNull).Index(1).Default(2312.231M);
                Map(x => x.TestData).Index(2);
            }
        }

        internal class TestClassMapDuplicateIndexes : ClassMap<TestClass>
        {
            public TestClassMapDuplicateIndexes()
            {
                Map(x => x.TestData).Index(1);
                Map(x => x.TestDataNull).Index(1);
            }
        }

        internal class TestClassMapDuplicateReadProperties : ClassMap<TestClass>
        {
            public TestClassMapDuplicateReadProperties()
            {
                Map(x => x.TestData).Index(1);
                Map(x => x.TestData).Index(2);
            }
        }

        internal class TestClassMapIndexOutOfRange : ClassMap<TestClass>
        {
            public TestClassMapIndexOutOfRange()
            {
                Map(x => x.TestData).Index(1);
                Map(x => x.TestDataNull).Index(16385);
            }
        }

        internal class TestClassMapLongHeaders : ClassMap<TestClass>
        {
            public TestClassMapLongHeaders()
            {
                string longHeader = new string('a', 256);
                Map(x => x.TestData).Index(1).Name(longHeader);
            }
        }

        internal class TestClassMapMissingIndexes : ClassMap<TestClass>
        {
            public TestClassMapMissingIndexes()
            {
                Map(x => x.TestData).Index(1);
                Map(x => x.TestDataNull);
            }
        }
    }
}