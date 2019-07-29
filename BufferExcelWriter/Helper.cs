using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace BufferExcelWriter
{
    public static class ExcelExportHelper
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = string.Concat(Convert.ToChar(65 + modulo), columnName);
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string FilterControlChar(string str)
        {
            return new string(str.Where(s => !char.IsControl(s)).ToArray());
        }

        public static void ExtractToDirectory(this ZipArchive archive, string destinationFullDirectoryName,
            bool overwrite)
        {
            if (!overwrite)
            {
                archive.ExtractToDirectory(destinationFullDirectoryName);
                return;
            }


            foreach (var file in archive.Entries)
            {
                var completeFileName = Path.GetFullPath(Path.Combine(destinationFullDirectoryName, file.FullName));

                if (!completeFileName.StartsWith(destinationFullDirectoryName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new IOException(
                        "Trying to extract file outside of destination directory. See this link for more info: https://snyk.io/research/zip-slip-vulnerability");
                }

                if (file.Name == "")
                {
                    // Assuming Empty for Directory
                    Directory.CreateDirectory(Path.GetDirectoryName(completeFileName) ??
                                              throw new Exception("NoDirectoryName"));
                    continue;
                }

                file.ExtractToFile(completeFileName, true);
            }
        }
    }
}