# BufferExcelWriter

1.  wb = new WorkBookDfn()  
2.  sheet = new WorkSheetDfn()  
    2.1. Add header(cells without value) to sheet's header
3.  Fill data to sheet.BufferRows  
4.  wb.FlushBufferedRowsAsync()
5.  repeat 3 when need
6.  add sheets to workbook
7.  BuildExcelAndGetStreamAsync (you can update sheet reference, sheet header, before call this)
8.  call Dispose


Samples Project => BufferExcelWriter.Sample
