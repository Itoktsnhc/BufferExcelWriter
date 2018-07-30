using System;
using System.Collections.Generic;
using System.Text;

namespace BufferExcelWriter
{
    public static class ExcelExportHelper
    {
        public static String GenerateFilterXmlEle(Int32 columnCount)
        {
            return $@"<autoFilter ref=""A1:{GetExcelColumnName(columnCount)}1"" xr:uid=""{Guid.NewGuid().ToString("B").ToUpper()}""/>";
        }

        public static String GetExcelColumnName(Int32 columnNumber)
        {
            var dividend = columnNumber;
            var columnName = String.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = String.Concat(Convert.ToChar(65 + modulo), columnName);
                dividend = (Int32)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}
