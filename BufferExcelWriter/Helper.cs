using System;
using System.Text.RegularExpressions;

namespace BufferExcelWriter
{
    public static class ExcelExportHelper
    {
        private static readonly Regex _pattern = new Regex("[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u10000-\u10FFF]+", RegexOptions.Compiled);

        public static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = String.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = String.Concat(Convert.ToChar(65 + modulo), columnName);
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string FilterOddChar(string str)
        {
            return _pattern.Replace(str, "");
        }
    }
}