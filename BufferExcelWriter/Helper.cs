using System;
using System.Linq;
using System.Text;
using System.Xml;

namespace BufferExcelWriter
{
    public static class ExcelExportHelper
    {
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

        public static string FilterControlChar(string str)
        {
            return new string(str.Where(s => !Char.IsControl(s)).ToArray());
        }

        public static string StripNonValidXMLCharacters(string textIn)
        {
            if (String.IsNullOrEmpty(textIn))
            {
                return textIn;
            }

            var textOut = new StringBuilder(textIn.Length);

            foreach (var current in textIn)
            {
                if ((current == 0x9 || current == 0xA || current == 0xB || current == 0xD) ||
                 ((current >= 0x20) && (current <= 0xD7FF)) ||
                 ((current >= 0xE000) && (current <= 0xFFFD)) ||
                 ((current >= 0x10000) && (current <= 0x10FFFF)))
                {
                    textOut.Append(current);
                }
            }

            return textOut.ToString();
        }
    }
}