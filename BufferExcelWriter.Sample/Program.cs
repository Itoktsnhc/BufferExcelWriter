using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace BufferExcelWriter.Sample
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            //await WriteTestDataAsync1();
            await WriteTestDataAsync2();

        }


        public static async Task WriteTestDataAsync1()
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
            var wb = new WorkBookDfn();
            wb.Sheets.Add(oddSheet);
            wb.Sheets.Add(evenSheet);


            await wb.OpenWriteExcelAsync(); //open write
            var sw = new Stopwatch();
            foreach (var outerIndex in Enumerable.Range(0, 10))
            {
                sw.Reset();
                var size = 1000;
                foreach (var index in Enumerable.Range(outerIndex * size, size))
                    if (index % 2 == 0)
                        evenSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Name", $"fo&o{index}"),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                    else
                        oddSheet.BufferedRows.Add(new RowDfn
                        {
                            Cells = new List<CellDfn>
                            {
                                new CellDfn("Nam®e", $"f￥￥￥©$ \"oo{index}"),
                                new CellDfn("Index", index.ToString())
                            }
                        });
                sw.Start();
                await wb.FlushBufferedRowsAsync(true); //flushDataAndClean
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

        public static async Task WriteTestDataAsync2()
        {
            var wb = new WorkBookDfn();

            try
            {
                var dataList = JsonConvert.DeserializeObject<List<BackFlowDto>>(await File.ReadAllTextAsync(@"D:\Export\data.json"));

                var header = new RowDfn
                {
                    Cells = new List<CellDfn>()
                    {
                        new CellDfn("Id"),
                        new CellDfn("Url"),
                        new CellDfn("Name"),
                        new CellDfn("tags"),
                        new CellDfn("contenttype"),
                        new CellDfn("teamproject"),
                        new CellDfn("taskflag"),
                        new CellDfn("ESCount"),
                        new CellDfn("数据获取方式"),
                        new CellDfn("备注"),
                        new CellDfn(null)
                    }
                };
                var sheet = new WorkSheetDfn("balabala", header);
                wb.Sheets.Add(sheet);
                await wb.OpenWriteExcelAsync();
                foreach (var dto in dataList)
                {
                    sheet.BufferedRows.Add(new RowDfn()
                    {
                        Cells = new List<CellDfn>()
                        {
                            new CellDfn("Id",dto.Id?.ToString()),
                            new CellDfn("Url",dto.Url),
                            new CellDfn("Name",dto.Name),
                            new CellDfn("tags",dto.Tags),
                            new CellDfn("contenttype",dto.ContentType),
                            new CellDfn("teamproject",dto.TeamProject),
                            new CellDfn("taskflag",dto.TaskFlag),
                            new CellDfn("ESCount",dto.EsCount?.ToString()),
                            new CellDfn("数据获取方式",dto.DataAccessMethod),
                            new CellDfn("备注",dto.Remark),
                            new CellDfn(null,null)
                        }
                    });
                }

                await wb.FlushBufferedRowsAsync(true);

                using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
                {
                    using (var stream = await wb.CloseExcelAndGetStreamAsync()) //close write and get stream from finished job
                    {
                        stream.Position = 0;
                        stream.CopyTo(fs);
                    }
                }



            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {
                wb.Dispose();
            }


        }
    }


    class BackFlowDto
    {
        public int? Id { get; set; }
        public string Url { get; set; }
        public string Name { get; set; }
        public string Tags { get; set; }
        public string ContentType { get; set; }
        public string TeamProject { get; set; }
        public string TaskFlag { get; set; }
        public long? EsCount { get; set; }
        public string DataAccessMethod { get; set; } = "Url";
        public string Remark { get; set; }
        public string SchemaName { get; set; }
        public int RootSchemaId { get; set; }
    }

    internal class Person
    {
        public string Name { get; set; }
        public bool Gender { get; set; }
        public uint Age { get; set; }
        public DateTime Birthday { get; set; }
    }
}