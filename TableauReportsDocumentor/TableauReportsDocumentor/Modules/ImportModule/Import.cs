using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.IO.Compression;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    class Import
    {
        public XmlDocument ImportTableauWorkbooks(string filename)
        {
            XmlDocument doc = new XmlDocument();
            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");

            XPathDocument myXPathDoc = new XPathDocument(filename);

            if (filename.EndsWith(".twb"))
            {
                myXPathDoc = new XPathDocument(filename);
            }
            else if (filename.EndsWith(".twbx"))
            {
                FileStream appArchiveFileStream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                using (ZipArchive zipArchive = new ZipArchive(appArchiveFileStream, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry entry in zipArchive.Entries)
                    {
                        if (entry.FullName.EndsWith(".twb", StringComparison.OrdinalIgnoreCase))
                        {
                            using (Stream appManifestFileStream = appManifestEntry.Open())
                            {
                                myXPathDoc = new XPathDocument(appManifestFileStream);
                            }
                        }
                    }
                }
            }

            using (XmlWriter writer = doc.CreateNavigator().AppendChild())
            {
                myXslTrans.Transform(myXPathDoc, null, writer);
            }

            return doc;
        }

    }
}
