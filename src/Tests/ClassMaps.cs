namespace Tests
{
    using OpenSpreadsheet.Configuration;
    using Xunit;

    public class ClassMaps
    {
        [Fact]
        public void TestConstantsValidation()
        {
            var validator = new ConfigurationValidator<TestClass, TestClassMapConstants>();
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
        }

        [Fact]
        public void TestDefaultsValidation()
        {
            var validator = new ConfigurationValidator<TestClass, TestClassMapDefaults>();
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
        }

        [Fact]
        public void TestDuplicateIndexesValidation()
        {
            var classMap = new TestClassMapDuplicateIndexes();
            var validator = new ConfigurationValidator<TestClass, TestClassMapDuplicateIndexes>();
            validator.Validate();
            Assert.True(validator.HasErrors);
        }

        [Fact]
        public void TestDuplicateReadPropertiesValidation()
        {
            var classMap = new TestClassMapDuplicateReadProperties();
            var validator = new ConfigurationValidator<TestClass, TestClassMapDuplicateReadProperties>();
            validator.Validate();
            Assert.True(validator.HasErrors);
        }

        [Fact]
        public void TestIndexOutOfRangeValidation()
        {
            var classMap = new TestClassMapIndexOutOfRange();
            var validator = new ConfigurationValidator<TestClass, TestClassMapIndexOutOfRange>();
            validator.Validate();
            Assert.True(validator.Errors.Count == 2);
        }

        [Fact]
        public void TestLongHeadersValidation()
        {
            var classMap = new TestClassMapLongHeaders();
            var validator = new ConfigurationValidator<TestClass, TestClassMapLongHeaders>();
            validator.Validate();
            Assert.True(validator.Errors.Count == 1);
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
                base.Map(x => x.TestData).Index(1).Constant(2312.231M);
                base.Map(x => x.TestDataNull).Index(2);
            }
        }

        internal class TestClassMapDefaults : ClassMap<TestClass>
        {
            public TestClassMapDefaults()
            {
                base.Map(x => x.TestDataNull).Index(1).Default(2312.231M);
                base.Map(x => x.TestData).Index(2);
            }
        }

        internal class TestClassMapDuplicateIndexes : ClassMap<TestClass>
        {
            public TestClassMapDuplicateIndexes()
            {
                base.Map(x => x.TestData).Index(1);
                base.Map(x => x.TestDataNull).Index(1);
            }
        }

        internal class TestClassMapDuplicateReadProperties : ClassMap<TestClass>
        {
            public TestClassMapDuplicateReadProperties()
            {
                base.Map(x => x.TestData).Index(1);
                base.Map(x => x.TestData).Index(2);
            }
        }

        internal class TestClassMapIndexOutOfRange : ClassMap<TestClass>
        {
            public TestClassMapIndexOutOfRange()
            {
                base.Map(x => x.TestData).Index(1);
                base.Map(x => x.TestDataNull).Index(16385);
            }
        }

        internal class TestClassMapLongHeaders : ClassMap<TestClass>
        {
            public TestClassMapLongHeaders()
            {
                string longHeader = new string('a', 256);
                base.Map(x => x.TestData).Index(1).Name(longHeader);
            }
        }

        internal class TestClassMapMissingIndexes : ClassMap<TestClass>
        {
            public TestClassMapMissingIndexes()
            {
                base.Map(x => x.TestData).Index(1);
                base.Map(x => x.TestDataNull);
            }
        }
    }
}