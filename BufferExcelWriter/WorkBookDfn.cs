using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BufferExcelWriter
{
    public class WorkBookDfn : IDisposable
    {
        internal const String WorksheetDefaultHeaders =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\"> <sheetData>";

        internal const String WorksheetDefaultFooter = "</worksheet>";
        internal const String SheetDataDefaultFooter = "</sheetData>";

        private readonly Dictionary<Int32, Int32> _rowOffsetDic = new Dictionary<Int32, Int32>();
        internal String OutPutFilePath;
        internal Stream OutputStream;
        internal String WorkingFolder;

        public WorkBookDfn()
        {
            WorkingFolder = Guid.NewGuid().ToString("N");
            if (Directory.Exists(WorkingFolder))
            {
                var existDir = new DirectoryInfo(WorkingFolder);
                existDir.Delete(true);
            }

            Directory.CreateDirectory(WorkingFolder);
            OutPutFilePath = WorkingFolder + ".zip";
            if (File.Exists(OutPutFilePath)) File.Delete(OutPutFilePath);
            var assembly = Assembly.GetExecutingAssembly();
            using (var fs = assembly.GetManifestResourceStream("BufferExcelWriter.exceltemplate"))
            {
                if (fs == null) throw new FileNotFoundException("BufferExcelWriter.exceltemplate");
                var zipFile = new ZipArchive(fs);
                zipFile.ExtractToDirectory(WorkingFolder);
            }

            Sheets = new List<WorkSheetDfn>();
            FolderEntry = new FolderEntry(WorkingFolder);
        }

        internal FolderEntry FolderEntry { get; set; }
        public IList<WorkSheetDfn> Sheets { get; set; }

        /// <summary>
        ///clean temp folder and file
        /// </summary>
        public void Dispose()
        {
            foreach (var sheet in Sheets) sheet.StreamWriter.Dispose();
            if (File.Exists(OutPutFilePath)) File.Delete(OutPutFilePath);
            OutputStream.Dispose();
        }

        public async Task OpenWriteExcelAsync()
        {
            for (var i = 0; i < Sheets.Count; i++)
            {
                var currentSheet = Sheets[i];
                if (String.IsNullOrWhiteSpace(currentSheet.Name)) currentSheet.Name = $"Sheet{i + 1}";
                currentSheet.SheetNum = i + 1;
            }

            foreach (var sheet in Sheets)
            {
                _rowOffsetDic[sheet.SheetNum] = 1;

                #region Update [Content_Types].xml

                var contentTypeEntry = FolderEntry.GetEntry("[Content_Types].xml");
                if (contentTypeEntry == null) throw new FileNotFoundException("[Content_Types].xml");

                using (var contentTypeStream = contentTypeEntry.Open())
                {
                    using (var sr = new StreamReader(contentTypeStream))
                    {
                        var doc = new XmlDocument();
                        doc.LoadXml(await sr.ReadToEndAsync());
                        var types = doc.GetElementsByTagName("Types")
                            .Cast<XmlNode>()
                            .First();
                        if (doc.DocumentElement != null)
                        {
                            var element = doc.CreateElement("Override", doc.DocumentElement.NamespaceURI);
                            element.SetAttribute("PartName", $"/xl/worksheets/sheet{sheet.SheetNum}.xml");
                            element.SetAttribute("ContentType",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                            types.AppendChild(element);
                        }

                        contentTypeStream.Position = 0;
                        contentTypeStream.SetLength(0);
                        doc.Save(contentTypeStream);
                    }
                }

                #endregion


                #region Update xl/_rels/workbook.xml.rels

                var identifier = "rId";
                var relsEntry = FolderEntry.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null) throw new FileNotFoundException("xl/_rels/workbook.xml.rels");

                using (var relsStream = relsEntry.Open())
                {
                    using (var sr = new StreamReader(relsStream))
                    {
                        var result = await sr.ReadToEndAsync();
                        var doc = new XmlDocument();
                        doc.LoadXml(result);
                        var relationships = doc.GetElementsByTagName("Relationships")
                            .Cast<XmlNode>()
                            .First();
                        identifier += (relationships.ChildNodes.Count + 1).ToString();
                        if (doc.DocumentElement != null)
                        {
                            var element = doc.CreateElement("Relationship", doc.DocumentElement.NamespaceURI);
                            element.SetAttribute("Target", $"worksheets/sheet{sheet.SheetNum}.xml");
                            element.SetAttribute("Type",
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                            element.SetAttribute("Id", identifier);
                            relationships.AppendChild(element);
                        }

                        relsStream.Position = 0;
                        relsStream.SetLength(0);
                        doc.Save(relsStream);
                    }
                }

                #endregion


                #region Update xl/workbook.xml

                var workbookEntry = FolderEntry.GetEntry("xl/workbook.xml");
                if (workbookEntry == null) throw new FileNotFoundException("xl/workbook.xml");

                using (var workbookStream = workbookEntry.Open())
                {
                    using (var sr = new StreamReader(workbookStream))
                    {
                        var doc = new XmlDocument();
                        doc.LoadXml(await sr.ReadToEndAsync());
                        var tags = doc.GetElementsByTagName("sheets")
                            .Cast<XmlNode>()
                            .First();
                        if (doc.DocumentElement != null)
                        {
                            var element = doc.CreateElement("sheet", doc.DocumentElement.NamespaceURI);
                            tags.AppendChild(element);
                            var attr = doc.CreateAttribute("r", "id",
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            attr.Value = identifier;
                            element.Attributes.Append(attr);
                            element.SetAttribute("sheetId", sheet.SheetNum.ToString());
                            element.SetAttribute("name", sheet.Name);
                        }

                        workbookStream.Position = 0;
                        workbookStream.SetLength(0);
                        doc.Save(workbookStream);
                    }
                }

                #endregion


                #region Update docProps/app.xml

                var appEntry = FolderEntry.GetEntry("docProps/app.xml");
                if (appEntry == null) throw new FileNotFoundException("docProps/app.xml Not Found in FolderEntry");

                using (var appStream = appEntry.Open())
                {
                    using (var sr = new StreamReader(appStream))
                    {
                        var result = sr.ReadToEnd();
                        var xd = new XmlDocument();
                        xd.LoadXml(result);
                        var element = xd.CreateElement("vt:lpstr",
                            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                        element.InnerText = sheet.Name;
                        var tmpEle = xd
                            .GetElementsByTagName("vt:vector")
                            .Cast<XmlNode>()
                            .Single(x => x.Attributes != null && x.Attributes["baseType"].Value == "lpstr");
                        tmpEle.AppendChild(element);

                        if (tmpEle.Attributes != null)
                            tmpEle.Attributes["size"].Value =
                                (Convert.ToInt32(tmpEle.Attributes["size"].Value) + 1).ToString();

                        var tmp2 = xd.GetElementsByTagName("vt:i4")
                            .Cast<XmlNode>()
                            .Single();
                        var val = String.IsNullOrWhiteSpace(tmp2.InnerText) ? 1 : Convert.ToInt32(tmp2.InnerText) + 1;
                        tmp2.InnerText = val.ToString();
                        appStream.Position = 0;
                        appStream.SetLength(0);
                        xd.Save(appStream);
                    }
                }

                #endregion


                #region Init sheetN.xml

                var sheetEntry = FolderEntry.CreateEntry(sheet.GetEntryName());
                sheet.FileStream = sheetEntry.Open();
                sheet.StreamWriter = new StreamWriter(sheet.FileStream);
                //sheet.FileStream.Position = sheet.FileStream.Length;
                await sheet.StreamWriter.WriteAsync(WorksheetDefaultHeaders);
                if (sheet.Header == null) throw new Exception("Header mustn't be null!");
                await sheet.StreamWriter.WriteAsync(sheet.Header.ToXmlString(1, sheet.Header, sheet.NullValStr));
                _rowOffsetDic[sheet.SheetNum]++;

                #endregion
            }
        }

        /// <summary>
        ///     Flush bufferedRow data to disk and clean buffered row
        /// </summary>
        /// <returns></returns>
        public async Task FlushBufferedRowsAsync(Boolean needGc = false)
        {
            foreach (var sheet in Sheets)
            {
                var sheetEntry = FolderEntry.GetEntry(sheet.GetEntryName());
                if (sheetEntry == null) throw new Exception("Init needed!!");

                var sb = new StringBuilder();
                for (var rowIndex = 0; rowIndex < sheet.BufferedRows.Count(); rowIndex++)
                {
                    var currentRow = sheet.BufferedRows[rowIndex];
                    sb.Append(currentRow.ToXmlString(_rowOffsetDic[sheet.SheetNum], sheet.Header,
                        sheet.NullValStr));
                    _rowOffsetDic[sheet.SheetNum]++;
                }

                await sheet.StreamWriter.WriteAsync(sb.ToString());
                await sheet.StreamWriter.FlushAsync();
                sheet.BufferedRows.Clear();
            }

            if (needGc)
            {
                GC.Collect();
            }
        }

        /// <summary>
        ///     write end info to file and get file stream back
        /// </summary>
        /// <returns></returns>
        public async Task<Stream> CloseExcelAndGetStreamAsync()
        {
            foreach (var sheet in Sheets)
            {
                var sheetEntry = FolderEntry.GetEntry(sheet.GetEntryName());
                if (sheetEntry == null) throw new FileNotFoundException(sheet.GetEntryName());
                await sheet.StreamWriter.WriteAsync(SheetDataDefaultFooter);
                await sheet.StreamWriter.WriteAsync(WorksheetDefaultFooter);
                await sheet.StreamWriter.FlushAsync();
                sheet.StreamWriter.Dispose();
            }

            ZipFile.CreateFromDirectory(WorkingFolder, OutPutFilePath, CompressionLevel.Optimal, false);
            if (Directory.Exists(WorkingFolder)) Directory.Delete(WorkingFolder, true);
            OutputStream = File.Open(OutPutFilePath, FileMode.Open);
            return OutputStream;
        }
    }

    internal class FolderEntry
    {
        private readonly String _dirPath;

        internal FolderEntry(String dirPath)
        {
            _dirPath = dirPath;
        }

        internal FileEntry GetEntry(String filename)
        {
            return new FileEntry(new FileInfo(Path.Combine(_dirPath, filename)));
        }

        internal FileEntry CreateEntry(String filename)
        {
            return GetEntry(filename);
        }
    }

    internal class FileEntry
    {
        private readonly FileInfo _fileInfo;

        internal FileEntry(FileInfo fileInfo)
        {
            _fileInfo = fileInfo;
        }

        internal Stream Open()
        {
            return _fileInfo.Open(FileMode.OpenOrCreate);
        }
    }
}