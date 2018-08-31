using System;
using System.Collections.Generic;
using System.IO;

namespace BufferExcelWriter
{
    public class WorkSheetDfn
    {
        public WorkSheetDfn(String sheetName, RowDfn header, String nullValStr = "-")
        {
            Name = sheetName;
            Header = header;
            NullValStr = nullValStr;
            BufferedRows = new List<RowDfn>();
        }

        public IList<RowDfn> BufferedRows { get; set; }
        public RowDfn Header { get; set; }
        internal String NullValStr { get; set; }
        internal String Name { get; set; }
        internal Int32 SheetNum { get; set; }
        internal Stream FileStream { get; set; }
        internal StreamWriter StreamWriter { get; set; }

        internal String GetEntryName()
        {
            return $"xl/worksheets/sheet{SheetNum}.xml";
        }
    }
}