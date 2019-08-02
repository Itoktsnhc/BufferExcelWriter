using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
// ReSharper disable once AutoPropertyCanBeMadeGetOnly.Global
// ReSharper disable once MemberCanBePrivate.Global

namespace BufferExcelWriter
{
    public class WorkBookDfn : IDisposable
    {
        private const string WorksheetDefaultHeaders =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\"> <sheetData>";

        private const string WorksheetDefaultFooter = "</worksheet>";
        private const string SheetDataDefaultFooter = "</sheetData>";
        private const string WorkingTempFolderName = ".temp";
        private const string TempBodyFileName = "bodyContent{0}.tmp";
        private readonly string _outPutFilePath;
        private readonly string _tempFolder;
        private readonly string _workingFolder;
        private Stream _outputStream;

        public WorkBookDfn(string baseDir = null)
        {
            if (string.IsNullOrEmpty(baseDir))
            {
                baseDir = Environment.CurrentDirectory;
            }

            if (!Path.IsPathRooted(baseDir))
            {
                baseDir = Path.Combine(Environment.CurrentDirectory, baseDir);
            }

            _workingFolder = Path.Combine(baseDir, Guid.NewGuid().ToString("N"));
            _tempFolder = Path.Combine(_workingFolder, WorkingTempFolderName);
            if (Directory.Exists(_workingFolder))
            {
                var existDir = new DirectoryInfo(_workingFolder);
                existDir.Delete(true);
            }

            if (Directory.Exists(_tempFolder))
            {
                var existDir = new DirectoryInfo(_tempFolder);
                existDir.Delete(true);
            }

            Directory.CreateDirectory(_workingFolder);
            Directory.CreateDirectory(_tempFolder);
            _outPutFilePath = _workingFolder + ".zip";
            if (File.Exists(_outPutFilePath))
            {
                File.Delete(_outPutFilePath);
            }

            Sheets = new List<WorkSheetDfn>();
            FolderEntry = new FolderEntry(_workingFolder);
            TempFolderEntry = new FolderEntry(_tempFolder);
        }

        private FolderEntry FolderEntry { get; }
        private FolderEntry TempFolderEntry { get; }


        public IList<WorkSheetDfn> Sheets { get; set; }

        /// <summary>
        ///     clean temp folder and file
        /// </summary>
        public void Dispose()
        {
            _outputStream?.Dispose();
            if (Directory.Exists(_workingFolder))
            {
                Directory.Delete(_workingFolder);
            }

            if (File.Exists(_outPutFilePath))
            {
                File.Delete(_outPutFilePath);
            }
        }

        private void InitFromZipFile()
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var fs = assembly.GetManifestResourceStream("BufferExcelWriter.exceltemplate"))
            {
                if (fs == null)
                {
                    throw new FileNotFoundException("BufferExcelWriter.ExcelTemplate");
                }

                var zipFile = new ZipArchive(fs);
                zipFile.ExtractToDirectory(_workingFolder, true);
            }
        }

        private async Task UpdateSheetRelationshipAsync()
        {
            InitFromZipFile();
            UpdateSheetNum();

            foreach (var sheet in Sheets)
            {
                #region Update [Content_Types].xml

                var contentTypeEntry = FolderEntry.GetEntry("[Content_Types].xml");
                if (contentTypeEntry == null)
                {
                    throw new FileNotFoundException("[Content_Types].xml");
                }

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
                if (relsEntry == null)
                {
                    throw new FileNotFoundException("xl/_rels/workbook.xml.rels");
                }

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
                if (workbookEntry == null)
                {
                    throw new FileNotFoundException("xl/workbook.xml");
                }

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
                if (appEntry == null)
                {
                    throw new FileNotFoundException("docProps/app.xml Not Found in FolderEntry");
                }

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
                        {
                            tmpEle.Attributes["size"].Value =
                                (Convert.ToInt32(tmpEle.Attributes["size"].Value) + 1).ToString();
                        }

                        var tmp2 = xd.GetElementsByTagName("vt:i4")
                            .Cast<XmlNode>()
                            .Single();
                        var val = string.IsNullOrWhiteSpace(tmp2.InnerText) ? 1 : Convert.ToInt32(tmp2.InnerText) + 1;
                        tmp2.InnerText = val.ToString();
                        appStream.Position = 0;
                        appStream.SetLength(0);
                        xd.Save(appStream);
                    }
                }

                #endregion
            }
        }

        private void UpdateSheetNum()
        {
            for (var i = 0; i < Sheets.Count; i++)
            {
                var currentSheet = Sheets[i];
                if (string.IsNullOrWhiteSpace(currentSheet.Name))
                {
                    currentSheet.Name = $"Sheet{i + 1}";
                }

                currentSheet.SheetNum = i + 1;
            }
        }

        /// <summary>
        ///     Flush bufferedRow data to disk and clean buffered row
        /// </summary>
        /// <returns></returns>
        public async Task FlushBufferedRowsAsync(bool needGc = false)
        {
            UpdateSheetNum();
            foreach (var sheet in Sheets)
            {
                var tempDataEntry = TempFolderEntry.CreateEntry(string.Format(TempBodyFileName, sheet.SheetNum));
                if (sheet.TempDataStreamWriter == null)
                {
                    sheet.TempDataStream = tempDataEntry.Open();
                    sheet.TempDataStreamWriter = new StreamWriter(sheet.TempDataStream);
                }

                var sheetEntry = FolderEntry.GetEntry(sheet.GetEntryName());
                if (sheetEntry == null)
                {
                    throw new Exception("Init needed!!");
                }

                foreach (var currentRow in sheet.BufferedRows)
                {
                    await sheet.TempDataStreamWriter.WriteLineAsync(JsonConvert.SerializeObject(currentRow));
                }

                await sheet.TempDataStreamWriter.FlushAsync();
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
        public async Task<Stream> BuildExcelAndGetStreamAsync(bool needGc = true)
        {
            if (Sheets.Count == 0)
            {
                throw new InvalidOperationException("no sheet in workbook");
            }

            await FlushBufferedRowsAsync(needGc);
            await UpdateSheetRelationshipAsync();

            foreach (var sheet in Sheets)
            {
                await GenerateSheetFileAsync(sheet);
            }

            Directory.Delete(_tempFolder, true);
            ZipFile.CreateFromDirectory(_workingFolder, _outPutFilePath, CompressionLevel.Optimal, false);
            if (Directory.Exists(_workingFolder))
            {
                Directory.Delete(_workingFolder, true);
            }

            _outputStream = File.Open(_outPutFilePath, FileMode.Open);
            return _outputStream;
        }

        private async Task GenerateSheetFileAsync(WorkSheetDfn sheet)
        {
            var sheetEntry = FolderEntry.GetEntry(sheet.GetEntryName());
            if (sheetEntry == null)
            {
                throw new FileNotFoundException(sheet.GetEntryName());
            }

            sheet.SheetFileStream = sheetEntry.Open();
            using (sheet.SheetStreamWriter = new StreamWriter(sheet.SheetFileStream))
            {
                await sheet.SheetStreamWriter.WriteAsync(WorksheetDefaultHeaders);
                var rowNumber = 1;
                await sheet.SheetStreamWriter.WriteAsync(sheet.Header.ToXmlString(rowNumber, sheet.Header,
                    sheet.NullValStr));
                if (sheet.TempDataStream != null)
                {
                    sheet.TempDataStream.Seek(0, SeekOrigin.Begin);
                    using (var streamReader = new StreamReader(sheet.TempDataStream))
                    {
                        while (!streamReader.EndOfStream)
                        {
                            rowNumber++;
                            var line = await streamReader.ReadLineAsync();
                            var currentRow = JsonConvert.DeserializeObject<RowDfn>(line);
                            await sheet.SheetStreamWriter.WriteAsync(currentRow.ToXmlString(rowNumber,
                                sheet.Header,
                                sheet.NullValStr));
                        }
                    }
                }

                await sheet.SheetStreamWriter.WriteAsync(SheetDataDefaultFooter);
                await sheet.SheetStreamWriter.WriteAsync(WorksheetDefaultFooter);
                await sheet.SheetStreamWriter.FlushAsync();
            }
        }
    }

    internal class FolderEntry
    {
        private readonly string _dirPath;

        internal FolderEntry(string dirPath)
        {
            _dirPath = dirPath;
        }

        internal FileEntry GetEntry(string filename)
        {
            return new FileEntry(new FileInfo(Path.Combine(_dirPath, filename)));
        }

        internal FileEntry CreateEntry(string filename)
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