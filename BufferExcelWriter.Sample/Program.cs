using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace BufferExcelWriter.Sample
{
    internal class Program
    {
        private static void Main(String[] args)
        {
            var header = new RowDfn
            {
                Cells = new List<CellDfn>
                {
                    new CellDfn("Name"),
                    new CellDfn("Index"),
                    new CellDfn("noVal")
                }
            };
            var oddSheet = new WorkSheetDfn("Odd", header); //add headers
            var evenSheet = new WorkSheetDfn("Even", header);
            var wb = new WorkBookDfn();
            wb.Sheets.Add(oddSheet);
            wb.Sheets.Add(evenSheet);


            wb.OpenWriteExcelAsync().Wait(); //open write
            var sw = new Stopwatch();
            foreach (var outerIndex in Enumerable.Range(0, 100))
            {
                sw.Reset();
                var size = 10000;
                foreach (var index in Enumerable.Range(outerIndex * size, size))
                    if (index % 2 == 0)
                        evenSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Name", $"foo{index}"),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                    else
                        oddSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Name", $"foo{index}"),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                sw.Start();
                wb.FlushBufferedRowsAsync(true).Wait(); //flushDataAndClean
                sw.Stop();
                Console.WriteLine(sw.Elapsed);
            }

            using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
            {
                using (var stream = wb.CloseExcelAndGetStreamAsync().Result)
                {
                    stream.Position = 0;
                    stream.CopyTo(fs);
                }
            }

            sw.Reset();
            sw.Start();
            wb.Dispose();
            sw.Stop();
            Console.WriteLine($"clean cost : {sw.Elapsed}");
            Console.WriteLine("Over");
            Console.ReadLine();
        }
    }

    internal class Person
    {
        public String Name { get; set; }
        public Boolean Gender { get; set; }
        public UInt32 Age { get; set; }
        public DateTime Birthday { get; set; }
    }
}