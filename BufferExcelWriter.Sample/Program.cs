using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BufferExcelWriter.Sample
{
    public static class Program
    {
        private static async Task Main()
        {
            await WriteTestDataAsyncSample0();
            await WriteTestDataAsyncSample1();
        }

        private static async Task WriteTestDataAsyncSample0()
        {
            var header = new RowDfn
            {
                Cells = new List<CellDfn>
                {
                    new CellDfn("Name"),
                    new CellDfn("Val"),
                    new CellDfn("noVal") //no value cell 
                }
            };
            //add headers
            var sheet = new WorkSheetDfn("Odd", header);

            sheet.BufferedRows.Add(new RowDfn
            {
                Cells = new List<CellDfn>
                {
                    new CellDfn("Name", "Hello"),
                    new CellDfn("Val", "World")
                }
            });

            //output
            using (var wb = new WorkBookDfn("tempFolder"))
            {
                wb.Sheets.Add(sheet);
                await wb.FlushBufferedRowsAsync(true);
                using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
                {
                    using (var stream = wb.BuildExcelAndGetStreamAsync().Result)
                    {
                        stream.Position = 0;
                        stream.CopyTo(fs);
                    }
                }
            }
        }

        private static async Task WriteTestDataAsyncSample1()
        {
            var header = new RowDfn
            {
                Cells = new List<CellDfn>
                {
                    new CellDfn("Nam®e"),
                    new CellDfn("Index"),
                    new CellDfn("noVal")
                }
            };
            var oddSheet = new WorkSheetDfn("Odd", header); //add headers
            var evenSheet = new WorkSheetDfn("Even", header);

            var wb = new WorkBookDfn("tempFolder");
            wb.Sheets.Add(oddSheet);
            wb.Sheets.Add(evenSheet);


            var sw = new Stopwatch();
            foreach (var outerIndex in Enumerable.Range(0, 10))
            {
                sw.Reset();
                var size = 10000;
                var strad = (char) 0xb;
                var inValidStr = new string(new[] {strad});
                foreach (var index in Enumerable.Range(outerIndex * size, size))
                {
                    if (index % 2 == 0)
                    {
                        evenSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Name", inValidStr),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                    }
                    else
                    {
                        oddSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Nam®e", $"f￥￥￥©$ \"oo{index}"),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                    }
                }


                sw.Start();
                await wb.FlushBufferedRowsAsync(true); //flushDataAndClean
                sw.Stop();
                Console.WriteLine(sw.Elapsed);
            }

            var insertSheet = new WorkSheetDfn("aaa", header);
            wb.Sheets.Add(insertSheet);
            evenSheet.Header.Cells.Add(new CellDfn("Latter1"));
            oddSheet.Header.Cells.Add(new CellDfn("Latter2"));
            sw.Reset();
            sw.Start();
            using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
            {
                using (var stream = wb.BuildExcelAndGetStreamAsync().Result)
                {
                    stream.Position = 0;
                    stream.CopyTo(fs);
                }
            }

            wb.Dispose();
            sw.Stop();
            Console.WriteLine($"build and clean cost : {sw.Elapsed}");
            Console.WriteLine("Over");
            Console.ReadLine();
        }
    }
}