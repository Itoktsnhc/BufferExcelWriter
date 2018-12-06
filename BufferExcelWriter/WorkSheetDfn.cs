using System.Collections.Generic;
using System.IO;

namespace BufferExcelWriter
{
    public class WorkSheetDfn
    {
        public WorkSheetDfn(string sheetName, RowDfn header, string nullValStr = "-")
        {
            Name = sheetName;
            Header = header;
            NullValStr = nullValStr;
            BufferedRows = new List<RowDfn>();
        }

        public IList<RowDfn> BufferedRows { get; set; }
        public RowDfn Header { get; set; }
        internal string NullValStr { get; set; }
        internal string Name { get; set; }
        internal int SheetNum { get; set; }
        internal Stream FileStream { get; set; }
        internal StreamWriter StreamWriter { get; set; }

        internal string GetEntryName()
        {
            return $"xl/worksheets/sheet{SheetNum}.xml";
        }
    }
}