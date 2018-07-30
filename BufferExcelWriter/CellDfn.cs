using System;
using System.Collections.Generic;
using System.Text;

namespace BufferExcelWriter
{
    public class CellDfn
    {
        /// <summary>
        /// create normalCell
        /// </summary>
        /// <param name="columnHeaderName"></param>
        /// <param name="value"></param>
        public CellDfn(String columnHeaderName, String value)
        {
            ColumnHeaderName = columnHeaderName;
            CellValue = value;
        }
        /// <summary>
        /// create headerCell
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
            return $"<c r=\"{ExcelExportHelper.GetExcelColumnName(columnNumber)}{rowNumber}\" t=\"inlineStr\"><is><t>{CellValue}</t></is></c>";
        }
    }

}
