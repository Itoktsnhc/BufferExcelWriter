using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection.Metadata;

namespace BufferExcelWriter.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            var header = new RowDfn()
            {
                Cells = new List<CellDfn>()
                    {
                        new CellDfn("Name"),
                        new CellDfn("Index"),
                        new CellDfn("noVal")
                    }
            };
            var oddSheet = new WorkSheetDfn("Odd", header);//add headers
            var evenSheet = new WorkSheetDfn("Even", header);
            var wb = new WorkBookDfn();
            wb.Sheets.Add(oddSheet);
            wb.Sheets.Add(evenSheet);


            wb.OpenWriteExcelAsync().Wait();//open write

            foreach (var outerIndex in Enumerable.Range(0, 100))
            {
                var size = 10000;
                foreach (var index in Enumerable.Range(outerIndex * size, size))
                {
                    if (index % 2 == 0)
                    {
                        evenSheet.BufferedRows.Add(new RowDfn()
                        {
                            Cells = new List<CellDfn>()
                                {
                                    new CellDfn("Name",$"foo{index}"),
                                    new CellDfn("Index",index.ToString())
                                }
                        });
                    }
                    else
                    {

                        oddSheet.BufferedRows.Add(new RowDfn()
                        {
                            Cells = new List<CellDfn>()
                                {
                                    new CellDfn("Name",$"foo{index}"),
                                    new CellDfn("Index",index.ToString())
                                }
                        });
                    }
                }

                wb.FlushBufferRowsAsync().Wait();//flushDataAndClean
                wb.CleanSheetsBuffer();
            }

            using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
            {
                using (var stream = wb.CloseExcelAndGetStreamAsync().Result)
                {
                    stream.Position = 0;
                    stream.CopyTo(fs);
                }
            }
            wb.Dispose();
            Console.WriteLine("Over");
            Console.ReadLine();
        }
    }

    class Person
    {
        public String Name { get; set; }
        public Boolean Gender { get; set; }
        public UInt32 Age { get; set; }
        public DateTime Birthday { get; set; }
    }
}
