using System.Collections.Generic;
using System.IO;
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Global

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
        public string Name { get; set; }
        internal int SheetNum { get; set; }
        internal Stream SheetFileStream { get; set; }
        internal Stream TempDataStream { get; set; }
        internal StreamWriter SheetStreamWriter { get; set; }
        internal StreamWriter TempDataStreamWriter { get; set; }


        internal string GetEntryName()
        {
            return $"xl/worksheets/sheet{SheetNum}.xml";
        }
    }
}