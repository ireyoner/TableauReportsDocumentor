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

//using System.IO.Compression.FileSystem;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    class ImportTWBandTWBX : ImportInterface
    {
        private XmlDocument ImportTWB(string filename)
        {
            if (filename.EndsWith(".twb"))
            {
                return ImportTWBfromXPathDocument(new XPathDocument(filename));
            }
            return null;
        }

        private XmlDocument ImportTWBfromXPathDocument(XPathDocument myXPathDoc)
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
                using (ZipArchive zipArchive = new ZipArchive(appArchiveFileStream, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry entry in zipArchive.Entries)
                    {
                        if (entry.FullName.EndsWith(".twb", StringComparison.OrdinalIgnoreCase))
                        {
                            using (Stream appManifestFileStream = entry.Open())
                            {
                                return ImportTWBfromXPathDocument(new XPathDocument(appManifestFileStream));
                            }
                        }
                    }
                }
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

        public string MenuItemText
        {
            get { return "twb & twbx"; }
        }

        public System.Windows.Media.Imaging.BitmapFrame MenuItemIcone
        {
            get { return null; }
        }

        public XmlDocument Import()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook|*.twb;*.twbx|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                return ImportTableauWorkbook(openFileDialog.FileName);
            }
            return null;
        }
    }
}
