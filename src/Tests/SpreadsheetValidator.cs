namespace Tests
{
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;

    public class SpreadsheetValidator
    {
        public IList<ValidationErrorInfo> Errors { get; private set; } = new List<ValidationErrorInfo>();

        public bool HasErrors => this.Errors.Count > 0;

        public void Validate(string spreadsheetFile)
        {
            using (var spreadsheet = SpreadsheetDocument.Open(spreadsheetFile, false))
            {
                var validator = new OpenXmlValidator();
                this.Errors = validator.Validate(spreadsheet).ToList();
            }
        }
    }
}