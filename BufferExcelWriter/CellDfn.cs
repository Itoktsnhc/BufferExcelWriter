// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Local
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Global
namespace BufferExcelWriter
{
    public class CellDfn
    {
        /// <summary>
        ///     create normalCell
        /// </summary>
        /// <param name="columnHeaderName"></param>
        /// <param name="value"></param>
        public CellDfn(string columnHeaderName, string value)
        {
            ColumnHeaderName = columnHeaderName;
            CellValue = value;
        }

        /// <summary>
        ///     create headerCell
        /// </summary>
        /// <param name="headerName"></param>
        public CellDfn(string headerName)
        {
            ColumnHeaderName = headerName;
            CellValue = headerName;
        }

        /// <summary>
        /// for json deserialize
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        internal CellDfn()
        {
        }

        public string ColumnHeaderName { get; set; }
        private string CellValue { get; set; }


        internal string ToXmlString(int rowNumber, int columnNumber, string nullValSymbol = "-")
        {
            var cellValue = CellValue;
            if (string.IsNullOrEmpty(cellValue))
            {
                cellValue = nullValSymbol;
            }

            var strVal = ExcelExportHelper.FilterControlChar(cellValue.Replace("]]>", "]]&gt;"));
            if (strVal.Length > 32766)
            {
                strVal = strVal.Substring(0, 32766);
            }

            return
                $"<c r=\"{ExcelExportHelper.GetExcelColumnName(columnNumber)}{rowNumber}\" t=\"inlineStr\"><is><t><![CDATA[{strVal}]]></t></is></c>";
        }
    }
}