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
    class ImportTWBandTWBX
    {
        private XmlDocument ImportTWB(string filename)
        {
            if (filename.EndsWith(".twb"))
            {
                return ImportTWB(new XPathDocument(filename));
            }
            return null;
        }

        private XmlDocument ImportTWB(XPathDocument myXPathDoc)
        {
            XmlDocument doc = new XmlDocument();

            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");

            using (XmlWriter writer = doc.CreateNavigator().AppendChild())
            {
                myXslTrans.Transform(myXPathDoc, null, writer);
            }
            return doc;
        }

        private XmlDocument ImportTWBX(string filename)
        {
            if (filename.EndsWith(".twbx"))
            {
                FileStream appArchiveFileStream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                //using (ZipArchive zipArchive = new ZipArchive(appArchiveFileStream, ZipArchiveMode.Read))
                //{
                //    foreach (ZipArchiveEntry entry in zipArchive.Entries)
                //    {
                //        if (entry.FullName.EndsWith(".twb", StringComparison.OrdinalIgnoreCase))
                //        {
                //            using (Stream appManifestFileStream = appManifestEntry.Open())
                //            {
                //                return ImportTWB(new XPathDocument(appManifestFileStream));
                //            }
                //        }
                //    }
                //}
            }
            return null;
        }

        public XmlDocument ImportTableauWorkbook(string filename)
        {
            if (filename.EndsWith(".twbx"))
            {
                return ImportTWBX(filename);
            }
            else if (filename.EndsWith(".twb"))
            {
                return ImportTWB(filename);
            }
            return null;
        }
    }
}
