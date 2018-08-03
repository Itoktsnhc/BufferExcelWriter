# BufferExcelWriter


Samples in BufferExcelWriter.Sample

```CSharp
    var wb = new WorkBookDfn();//new workbook
    var header = new RowDfn//create header
    {
        Cells = new List<CellDfn>
        {
            new CellDfn("Name"),
            new CellDfn("Index"),
            new CellDfn("noVal")
        }
    };
    var sheet=new WorkSheetDfn("sheetName", header);//new sheet
    wb.Sheets.Add(sheet);//add sheet to workbook

    await wb.OpenWriteExcelAsync();//init write;
    /*
    balabala generate data like: */
    foreach (var outerIndex in Enumerable.Range(0, 100))
    {
         foreach (var index in Enumerable.Range(outerIndex * size, size))
         {
             
            sheet.BufferedRows.Add(new RowDfn
            {
                Cells = new List<CellDfn>
                {
                    new CellDfn("Name", $"foo{index}"),
                    new CellDfn("Index", index.ToString())
                }
            });
         }
         wb.FlushBufferedRowsAsync(true);//flush buffered row and clean buffered row
    }
   

    using (var fs = File.Create($"{DateTime.Now.Ticks}.xlsx"))
    {
        using (var stream = wb.CloseExcelAndGetStreamAsync().Result)//close write and get stream from finished job
        {
            stream.Position = 0;
            stream.CopyTo(fs);
        }
    }
    wb.Dispose();//clean stream、files and something else;

```
