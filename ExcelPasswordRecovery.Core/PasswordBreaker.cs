using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelPasswordRecovery.Core
{
    public class PasswordBreaker
    {
        public List<PasswordEntry> GetPasswords(string file)
        {
            if (!File.Exists(file))
            {
                throw  new ArgumentException("Cannot find file '{0}'", file);
            }

            var passwordEntries = new List<PasswordEntry>();

            using (var s = new ZipInputStream(File.OpenRead(file))) {

                ZipEntry entry;
                while ((entry = s.GetNextEntry()) != null)
                {
                    var path = Path.GetDirectoryName(entry.Name);
                    var fileName = Path.GetFileName(entry.Name);

                    if (path == "xl" && fileName == "workbook.xml")
                    {
                        // not implemented yet
                    }

                    if (path == "xl\\worksheets")
                    {
                        var document = ReadEntry(s);

                        var sheetProtection = document.GetElementsByTagName("sheetProtection");

                        var password = ParseEntries(sheetProtection);

                        if(password == null)
                            continue;
                        
                        passwordEntries.Add(new PasswordEntry
                        {
                            Name = "name",
                            Id = "id",
                            Type = PasswordEntryType.Worksheet
                        });

                    }
                }
            }

            return passwordEntries;
        }

        private static XmlDocument ReadEntry(ZipInputStream input)
        {
            var document = new XmlDocument();
            document.LoadXml(new StreamReader(input).ReadToEnd());

            return document;
        }

        private static string ParseEntries(XmlNodeList entries)
        {
            if(entries == null || entries.Count == 0)
                return null;

            var attributes = entries[0].Attributes;

            if (attributes == null)
                return null;
            
            var hash = attributes["password"].Value;

            // Get the password
            return ExcelHashUtility.FindPassword(hash);
        }
    }
}
