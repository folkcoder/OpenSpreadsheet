namespace OpenSpreadsheet.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    public static class TypeExtensions
    {
        private static readonly HashSet<Type> numericTypes = new HashSet<Type>()
        {
            typeof(byte), typeof(decimal), typeof(double), typeof(float), typeof(int), typeof(long), typeof(sbyte), typeof(short), typeof(uint), typeof(ulong), typeof(ushort),
        };

        public static bool IsNumeric(this Type type) => numericTypes.Contains(type);
    }
}