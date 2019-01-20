# SpreadsheetHelper

SpreadsheetHelper is a fast and lightweight wrapper around the OpenXml spreadsheet library, employing an easy-to-use fluent interface to define relations between entities and spreadsheet rows. The library uses the Simple API for XML (SAX) method for both reading and writing.

The primary use case for SpreadsheetHelper is efficiently importing and exporting typed collections, where each row roughly corresponds to a class instance. It is not meant to offer fine-grained control of data or formatting at the cell level; if you need this level of control, check out [ClosedXml](https://github.com/ClosedXML/ClosedXML) or [EPPlus](https://github.com/JanKallman/EPPlus).


## Syntax

### Configuration

SpreadsheetHelper uses a fluent interface to map object properties to spreadsheet rows.

**Basic Example**

Classes being mapped must have a parameterless constructor. A basic ClassMap requires a property (usually) and an index. The name is optional; for writing, if no name is provided the mapped property's name will be used. 

Most configuration properties have both a read and write version, if applicable. If you need to a class to have different mappings for reading and writing operations, simply use the appropriate map method.

```c#
public class TestClassMap : ClassMap<TestClass>
{
    public TestClassMap()
    {
        Map(x => x.Surname).Index(1).Name("Employee Last Name");
        Map(x => x.GivenName).Index(2).Name("Employee First Name");
        Map(x => x.Id).Index(3).Name("Employee Id");
        Map(x => x.Address).Index(4).IgnoreWrite(true);
        Map(x => x.SSN).IndexRead(10).IndexWrite(5).CustomNumberFormat("000-00-0000");
        Map(x => x.Amount).Index(6).NumberFormat();
    }
}
````


**Constants and Defaults**
If you need to supply a constant value to a property during reading or you'd like to write a constant value (with or without an associated property), use the Constant map.

If you need to supply a fallback value for null values, use the Default map.

```c#
public class TestClassMap : ClassMap<TestClass>
{
    public TestClassMap()
    {
        Map().Index(1).Name("Date").ConstantWrite(DateTime.Today.ToString());
        Map(x => x.Id).Index(2).Name("Employee Id").Default(0);
    }
}
````


**Column Styles**

In order to customize the appearance of a style, simply create a new ColumnStyle instance and map it to the property using the Style method. If an explicit ColumnStyle is not specified, a default instance will be used.

```c#
public class TestClassMap : ClassMap<TestClass>
{
    public TestClassMap()
    {
        var columnStyle = new ColumnStyle()
        {
            BackgroundColor = Color.Aquamarine,
            BackgroundPatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid,
            BorderColor = Color.Red,
            BorderPlacement = BorderPlacement.Outside,
            BorderStyle = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
            Font = new Font("Arial", 14, FontStyle.Italic),
            ForegroundColor = Color.White,
            HoizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
            NumberFormat = OpenXmlNumberingFormat.Currency,
            VerticalAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center
        };

        Map(x => x.Amount).Index(1).Style(columnStyle);
        Map(x => x.SSN).Index(2).Style(new ColumnStyle() { CustomNumberFormat = "000-00-0000" });
    }
}
```

### Writing

To write data to a new worksheet, simply call the WriteWorksheet method from your Spreadsheet, providing the type of object to be written and its associtiated map. If you want more fine-grained control over the write operation, have your Spreadsheet create a new WorksheetWriter.

```c#
using (var spreadsheet = new Spreadsheet(filepath))
{
    // write all records from the Spreadsheet (uses a WorksheetWriter behind the scenes)
    spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet2", records);

    // write all records using an explicit WorksheetWriter
    using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("Sheet3"))
    {
        writer.WriteRecords(records);
    }

    // write individual records from the WorksheetWriter
    using (var writer = spreadsheet.CreateWorksheetWriter<TestClass, TestClassMap>("Sheet1", 0))
    {
        writer.WriteHeader();
        writer.SkipRows(3);
        writer.WriteRecord(new TestClass() { TestData = "first row" });
        writer.WriteRecord(new TestClass() { TestData = "second row" });        
        writer.WriteRecord(new TestClass() { TestData = "third row" });
        writer.SkipRow();
        writer.WriteRecord(new TestClass() { TestData = "fourth row" });
    }
}
```

To apply general worksheet styles, create a new WorksheetStyle instance and pass it as an argument to your write operations. Otherwise, a default WorksheetStyle instance will be used.

```c#
var worksheetStyle = new WorksheetStyle()
{
    HeaderBackgroundColor = Color.Chartreuse,
    HeaderBackgroundPatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid,
    HeaderFont = new Font("Comic Sans", 16, FontStyle.Strikeout),
    HeaderForegroundColor = Color.DarkBlue,
    HeaderHoizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center,
    HeaderRowIndex = 2,
    HeaderVerticalAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center,
    MaxColumnWidth = 30,
    MinColumnWidth = 10,
    ShouldAutoFilter = true,
    ShouldAutoFitColumns = true,
    ShouldFreezeTopRow = true,
    ShouldWriteHeaderRow = true,
};

using (var spreadsheet = new Spreadsheet(filepath))
{
    spreadsheet.WriteWorksheet<TestClass, TestClassMap>("Sheet1", records, worksheetStyle);
}
```

### Reading

To read data from an existing worksheet, simply call the ReadWorksheet method from your Spreadsheet, providing the type of object to be written and its associtiated map. If you want more fine-grained control over the read operation, have your Spreadsheet create a new WorksheetReader.

```c#
using (var spreadsheet = new Spreadsheet(filepath))
{
    // read all records from the Spreadsheet (uses a WorksheetReader behind the scenes)
    var recordsSheet1 = spreadsheet.ReadWorksheet<TestClass, TestClassMap>("Sheet1");

    // read all records using an explicit WorksheetReader
    using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>("Sheet2"))
    {
        var recordsSheet2 = reader.ReadRows();
    }

    // read individual records
    using (var reader = spreadsheet.CreateWorksheetReader<TestClass, TestClassMap>("Sheet3"))
    {
        var firstRow = reader.ReadRow();
        var secondRow = reader.ReadRow();
        reader.SkipRow();
        var fourthRow = reader.ReadRow();
    }
}
```

## Performance

**Reading**

SpreadsheetHelper has slightly better memory performance than ClosedXml and EPPlus, but runs slightly slower than EPPlus. For reading, all three libraries are pretty performant.

| Library | Records | Fields | Operation | Runtime | Memory Used |
| ------------- |
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 50,000 | 3 | 987.6 ms | 212.02 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 50,000 | 3 | 417.7 ms | 151.51 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 50,000 | 3 | 663.5 ms | 109.32 MB
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 100,000 | 3 | 1,981.3 ms | 424.68 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 100,000 | 3 | 833.3 ms | 302.11 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 100,000 | 3 | 1,323.9 ms | 217.72 MB
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 250,000 | 3 | 4,967.7 ms | 1046.61 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 250,000 | 3 | 2,082.8 ms | 744.05 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 250,000 | 3 | 3,278.5 ms | 541.94 MB


**Writing**

SpreadsheetHelper is significantly faster and more memory-friendly than ClosedXml, and slightly more so than EPPlus.

| Library | Records | Fields | Operation | Runtime | Memory Used |
| ------------- |
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 50,000 | 30 | 17.162 s | 212.05 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 50,000 | 30 | 3.877 s | 151.51 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 50,000 | 30 | 3.654 s | 109.32 MB
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 100,000 | 30 | 35.328 s | 6957.00 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 100,000 | 30 | 7.708 s | 2243.08 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 100,000 | 30 | 7.244 s | 1742.74 MB
| [ClosedXml](https://github.com/ClosedXML/ClosedXML) | 250,000 | 30 | 82.006 s | 17277.52 MB
| [EPPlus](https://github.com/JanKallman/EPPlus) | 250,000 | 30 | 19.753 s | 5514.32 MB
| [SpreadsheetHelper](https://github.com/FolkCoder/SpreadsheetHelper) | 250,000 | 30 | 18.524 s | 4340.02 MB




## Future Plans

+ Make index maps optional. For writing, unspecified indexes would default to map order, and read indexes would attempt to map column indexes by header name.
+ Make additional tweaks to configurtion maps to make them easier to use and validate.
+ Migrate library to .NET Core. A long-outstanding bug in the corefx library is preventing this transition (https://github.com/dotnet/corefx/issues/24457).
