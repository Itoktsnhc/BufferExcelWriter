using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Transactions;
using System.Xml;

namespace BufferExcelWriter
{
    public class WorkBookDfn
    {
        public ZipArchive ZipArchive { get; set; }
        public IList<WorkSheetDfn> Sheets { get; set; }

        public async Task OpenWriteExcelAsync()
        {
            for (int i = 0; i < Sheets.Count; i++)
            {
                var currentSheet = Sheets[i];
                if (String.IsNullOrWhiteSpace(currentSheet.Name))
                {
                    currentSheet.Name = $"Sheet{i + 1}";
                }
                currentSheet.SheetNum = i + 1;
            }

            foreach (var sheet in Sheets)
            {

                #region Update [Content_Types].xml

                var contentTypeEntry = ZipArchive.GetEntry("[Content_Types].xml");
                if (contentTypeEntry == null)
                {
                    throw new FileNotFoundException("[Content_Types].xml Not Found in ZipArchive");
                }

                using (var contentTypeStream = contentTypeEntry.Open())
                {
                    using (var sr = new StreamReader(contentTypeStream))
                    {
                        var doc = new XmlDocument();
                        doc.Load(await sr.ReadToEndAsync());
                        if (doc.DocumentElement != null)
                        {
                            var element = doc.CreateElement("Override", doc.DocumentElement.NamespaceURI);
                            element.SetAttribute("PartName", $"/xl/worksheets/sheet{sheet.SheetNum}.xml");
                            element.SetAttribute("ContentType",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                            contentTypeStream.Position = 0;
                            contentTypeStream.SetLength(0);
                            doc.Save(contentTypeStream);
                        }
                    }
                }

                #endregion


                #region Update xl/_rels/workbook.xml.rels

                string identifier = "rId";
                var relsEntry = ZipArchive.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null)
                {
                    throw new FileNotFoundException("xl/_rels/workbook.xml.rels Not Found in ZipArchive");
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

                var workbookEntry = ZipArchive.GetEntry("xl/workbook.xml Not Found in ZipArchive");
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

                var appEntry = ZipArchive.GetEntry("docProps/app.xml");
                if (appEntry == null)
                {
                    throw new FileNotFoundException("docProps/app.xml Not Found in ZipArchive");
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
                            tmpEle.Attributes["size"].Value =
                                (Convert.ToInt32(tmpEle.Attributes["size"].Value) + 1).ToString();

                        var tmp2 = xd.GetElementsByTagName("vt:i4")
                            .Cast<XmlNode>()
                            .Single();
                        tmp2.InnerText = (Convert.ToInt32(tmp2.InnerText) + 1).ToString();
                        appStream.Position = 0;
                        appStream.SetLength(0);
                        xd.Save(appStream);
                    }
                }

                #endregion


                #region Init sheetN.xml

                var sheetEntry = ZipArchive.CreateEntry(sheet.GetEntryName());
                using (var sheetStream = sheetEntry.Open())
                {
                    using (var sw = new StreamWriter(sheetStream))
                    {
                        await sw.WriteAsync(WorksheetDefaultHeaders);
                    }
                }

                #endregion
            }
        }

        public async Task AppendBufferRowsAsync()
        {
            foreach (var currentSheet in Sheets)
            {
                var sheetEntry = ZipArchive.GetEntry(currentSheet.GetEntryName());
                if (sheetEntry == null)
                {
                    throw new Exception("Init needed!!");
                }
                using (var sheetStream = sheetEntry.Open())
                {
                    using (var sw = new StreamWriter(sheetStream))
                    {

                        if (currentSheet.Header == null)
                        {
                            throw new Exception("Header mustn't be null!");
                        }

                        await sw.WriteAsync(currentSheet.Header.ToXmlString(1, currentSheet.Header, currentSheet.NullValStr));
                        for (var rowIndex = 0; rowIndex < currentSheet.BufferedRows.Count(); rowIndex++)
                        {
                            var currentRow = currentSheet.BufferedRows[rowIndex];
                            await sw.WriteAsync(currentRow.ToXmlString(rowIndex + 2, currentSheet.Header, currentSheet.NullValStr));
                        }
                        await sw.FlushAsync();
                    }
                }
            }

        }

        public async Task CloseWriteExcelAsync()
        {
            foreach (var sheet in Sheets)
            {
                var sheetEntry = ZipArchive.CreateEntry(sheet.GetEntryName());
                using (var sheetStream = sheetEntry.Open())
                {
                    using (var sw = new StreamWriter(sheetStream))
                    {
                        await sw.WriteAsync(SheetDataDefaultFooter);
                        await sw.WriteAsync(WorksheetDefaultFooter);
                        await sw.FlushAsync();
                    }
                }
            }
        }





        public const String WorksheetDefaultHeaders = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\"> <sheetData>";
        public const String WorksheetDefaultFooter = "</worksheet>";
        public const String SheetDataDefaultFooter = "</sheetData>";
    }
}
