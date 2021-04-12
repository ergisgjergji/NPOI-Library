using Npoi_Library.Excel.Enums;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Npoi_Library.Excel.Helpers
{
    public static class ExcelHelpers
    {
        /// <summary>
        /// Converts Excel column number to letter
        /// </summary>
        public static string ColNumberToLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        /// <summary>
        /// Converts Excel column letter to number
        /// </summary>
        public static int ColLetterToNumber(string colLetter)
        {
            string letter = colLetter.ToUpper();

            int[] digits = new int[letter.Length];
            for (int i = 0; i < letter.Length; ++i)
            {
                digits[i] = Convert.ToInt32(letter[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        public static FieldType GetFieldType(Type type)
        {
            IList<Type> numberTypeList = new List<Type> { typeof(byte), typeof(sbyte), typeof(short), typeof(ushort), typeof(int), typeof(uint), typeof(long), typeof(ulong), typeof(float), typeof(double), typeof(decimal), typeof(byte?), typeof(sbyte?), typeof(short?), typeof(ushort?), typeof(int?), typeof(uint?), typeof(long?), typeof(ulong?), typeof(float?), typeof(double?), typeof(decimal?) };

            if (type.Equals(typeof(string)) || type.Equals(typeof(char)) || type.Equals(typeof(char?)))
                return FieldType.Text;

            if (numberTypeList.Contains(type))
                return FieldType.Numeric;

            if (type.Equals(typeof(bool)) || type.Equals(typeof(bool?)))
                return FieldType.Boolean;

            if (type.Equals(typeof(DateTime)) || type.Equals(typeof(DateTime?)))
                return FieldType.DateTime;

            return FieldType.Other;
        }
    }
}
