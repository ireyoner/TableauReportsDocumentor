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
using TableauReportsDocumentor.Data;

//using System.IO.Compression.FileSystem;

namespace TableauReportsDocumentor.Modules.ImportModule
{
    class ImportTWBandTWBX : ImportInterface
    {
        private ImportedDocument ImportTWB(string filename)
        {
            if (filename.EndsWith(".twb"))
            {
                XmlDocument od= new XmlDocument();
                od.Load(filename);
                return new ImportedDocument(od,ImportTWBfromXPathDocument(new XPathDocument(filename)));
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

        private ImportedDocument ImportTWBX(string filename)
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
                                XmlDocument od = new XmlDocument();
                                od.Load(appManifestFileStream);
                                return new ImportedDocument(od, ImportTWBfromXPathDocument(new XPathDocument(appManifestFileStream)));
                            }
                        }
                    }
                }
            }
            return null;
        }

        public ImportedDocument ImportTableauWorkbook(string filename)
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

        public ImportedDocument Import()
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
