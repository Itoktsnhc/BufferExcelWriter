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
        private static void Main()
        {
            WriteTestDataSample1();
            WriteTestDataSampleAsync().Wait();
            Console.ReadLine();
        }

        private static void WriteTestDataSample1()
        {
            var sw = new Stopwatch();
            sw.Start();
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



            foreach (var outerIndex in Enumerable.Range(0, 10))
            {
                sw.Reset();
                var size = 10000;
                var strad = (char)0xb;
                var inValidStr = new string(new[] { strad });
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
                wb.FlushBufferedRows(true); //flushDataAndClean
            }

            var insertSheet = new WorkSheetDfn("aaa", header);
            wb.Sheets.Add(insertSheet);
            evenSheet.Header.Cells.Add(new CellDfn("Latter1"));
            oddSheet.Header.Cells.Add(new CellDfn("Latter2"));
            using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
            {
                using (var stream = wb.BuildExcelAndGetStream())
                {
                    stream.Position = 0;
                    stream.CopyTo(fs);
                }
            }

            wb.Dispose();
            sw.Stop();
            Console.WriteLine($"WriteTestDataSample1 cost : {sw.Elapsed}");
        }

        private static async Task WriteTestDataSampleAsync()
        {
            var sw = new Stopwatch();
            sw.Start();
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
            foreach (var outerIndex in Enumerable.Range(0, 10))
            {
                var size = 10000;
                var strad = (char)0xb;
                var inValidStr = new string(new[] { strad });
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
                await wb.FlushBufferedRowsAsync(true); //flushDataAndClean
            }

            var insertSheet = new WorkSheetDfn("aaa", header);
            wb.Sheets.Add(insertSheet);
            evenSheet.Header.Cells.Add(new CellDfn("Latter1"));
            oddSheet.Header.Cells.Add(new CellDfn("Latter2"));
            using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
            {
                using (var stream = wb.BuildExcelAndGetStream())
                {
                    stream.Position = 0;
                    stream.CopyTo(fs);
                }
            }

            wb.Dispose();
            sw.Stop();
            Console.WriteLine($"WriteTestDataSampleAsync cost : {sw.Elapsed}");
        }
    }
}