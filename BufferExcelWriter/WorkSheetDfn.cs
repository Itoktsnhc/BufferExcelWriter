using System;
using System.Collections.Generic;
using System.Text;

namespace BufferExcelWriter
{
    public class WorkSheetDfn
    {
        public WorkSheetDfn(String sheetName, RowDfn header, String nullValStr = "-")
        {
            Name = sheetName;
            Header = header;
            NullValStr = nullValStr;
        }
        internal IList<RowDfn> BufferedRows { get; set; }
        public RowDfn Header { get; set; }
        public String NullValStr { get; set; }
        public String Name { get; set; }
        internal Int32 SheetNum { get; set; }
        public String GetEntryName()
        {
            return $"xl/worksheets/sheet{SheetNum}.xml";
        }
    }
}
