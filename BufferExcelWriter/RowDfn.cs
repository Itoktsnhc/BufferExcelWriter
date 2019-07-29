using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BufferExcelWriter
{
    public class RowDfn
    {
        // ReSharper disable once MemberCanBePrivate.Global
        public IList<CellDfn> Cells { get; set; }

        internal string ToXmlString(int rowNumber, RowDfn header, string nullValSymbol = "-")
        {
            var row = new StringBuilder();
            if (Cells != null && Cells.Any())
            {
                row.Append($"<row r=\"{rowNumber}\" spans=\"{1}:{header.Cells.Count}\">");
                try
                {
                    for (var columnNumber = 0; columnNumber < header.Cells.Count; columnNumber++)
                    {
                        var headerCell = header.Cells[columnNumber];
                        var cell = Cells.FirstOrDefault(s => s.ColumnHeaderName == headerCell.ColumnHeaderName) ??
                                   new CellDfn(headerCell.ColumnHeaderName, nullValSymbol);

                        row.Append(cell.ToXmlString(rowNumber, columnNumber + 1, nullValSymbol));
                    }
                }
                finally
                {
                    row.Append("</row>");
                }
            }

            return row.ToString();
        }
    }
}