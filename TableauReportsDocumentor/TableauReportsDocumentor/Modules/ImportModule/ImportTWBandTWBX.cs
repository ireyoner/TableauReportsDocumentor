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
        private String original;
        public string OriginalReport
        {
            get { return original; }
        }

        private XmlDocument converted;
        public XmlDocument ConvertedReport
        {
            get { return converted; }
        }

        private Boolean ImportTWB(string filename)
        {
            if (filename.EndsWith(".twb"))
            {
                StreamReader streamReader = new StreamReader(filename);
                original = streamReader.ReadToEnd();
                streamReader.Close();
                return ImportTWBfromXPathDocument(new XPathDocument(filename));
            }
            return false;
        }

        private Boolean ImportTWBfromXPathDocument(XPathDocument myXPathDoc)
        {

            converted = new XmlDocument();

            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load("../../Import Converters/TRDCI_v8.3.xsl");

            using (XmlWriter writer = converted.CreateNavigator().AppendChild())
            {
                myXslTrans.Transform(myXPathDoc, null, writer);

            }
            return true;
        }

        private Boolean ImportTWBX(string filename)
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
                                StreamReader streamReader = new StreamReader(appManifestFileStream);
                                original = streamReader.ReadToEnd();
                                streamReader.Close();
                                return ImportTWBfromXPathDocument(new XPathDocument(appManifestFileStream));
                            }
                        }
                    }
                }
            }
            return false;
        }

        public Boolean ImportTableauWorkbook(string filename)
        {
            original = null;
            if (filename.EndsWith(".twbx"))
            {
                return ImportTWBX(filename);
            }
            else if (filename.EndsWith(".twb"))
            {
                return ImportTWB(filename);
            }
            return false;
        }

        public string MenuItemText
        {
            get { return "twb & twbx"; }
        }

        public System.Windows.Media.Imaging.BitmapFrame MenuItemIcone
        {
            get { return null; }
        }

        public Boolean Import()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook|*.twb;*.twbx|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                return ImportTableauWorkbook(openFileDialog.FileName);
            }
            return false;
        }



    }
}
