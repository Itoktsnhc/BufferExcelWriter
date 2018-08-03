using System;

namespace BufferExcelWriter
{
    public class CellDfn
    {
        /// <summary>
        ///     create normalCell
        /// </summary>
        /// <param name="columnHeaderName"></param>
        /// <param name="value"></param>
        public CellDfn(String columnHeaderName, String value)
        {
            ColumnHeaderName = columnHeaderName;
            CellValue = value;
        }

        /// <summary>
        ///     create headerCell
        /// </summary>
        /// <param name="headerName"></param>
        public CellDfn(String headerName)
        {
            ColumnHeaderName = headerName;
            CellValue = headerName;
        }

        public String ColumnHeaderName { get; set; }
        public String CellValue { get; set; }


        internal String ToXmlString(Int32 rowNumber, Int32 columnNumber)
        {
            var strVal = CellValue.Replace("]]>", "]]&gt;");
            if (strVal.Length > 32766) strVal = strVal.Substring(0, 32766);
            return
                $"<c r=\"{ExcelExportHelper.GetExcelColumnName(columnNumber)}{rowNumber}\" t=\"inlineStr\"><is><t>{strVal}</t></is></c>";
        }
    }
}