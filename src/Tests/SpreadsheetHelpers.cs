namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public static class SpreadsheetHelpers
    {
        public static uint GetColumnIndexFromCellReference(string cellReference)
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
    }
}