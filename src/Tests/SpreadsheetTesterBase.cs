namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    public abstract class SpreadsheetTesterBase
    {
        protected SpreadsheetValidator SpreadsheetValidator { get; } = new SpreadsheetValidator();

        public string ConstructTempExcelFilePath() => Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");

        public uint GetColumnIndexFromCellReference(string cellReference)
        {
            var columnLetters = cellReference.Where(c => !char.IsNumber(c)).ToArray();
            string columnReference = new string(columnLetters);

            int sum = 0;
            for (int i = 0; i < columnLetters.Length; i++)
            {
                sum *= 26;
                sum += columnLetters[i] - 'A' + 1;
            }

            return (uint)sum;
        }

        public IEnumerable<T> CreateRecords<T>(int count) where T : class
        {
            for (int i = 0; i < count; i++)
            {
                yield return Activator.CreateInstance<T>();
            }
        }
    }
}