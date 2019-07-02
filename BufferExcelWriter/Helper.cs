using System;
using System.Linq;

namespace BufferExcelWriter
{
    public static class ExcelExportHelper
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = string.Concat(Convert.ToChar(65 + modulo), columnName);
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string FilterControlChar(string str)
        {
            return new string(str.Where(s => !char.IsControl(s)).ToArray());
        }
    }
}