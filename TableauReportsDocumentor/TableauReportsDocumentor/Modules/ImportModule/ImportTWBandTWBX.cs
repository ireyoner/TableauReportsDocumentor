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
                var doc = new XmlDocument();
                doc.Load(filename);
                return ImportTWBfromXmlDocument(doc);
            }
            return false;
        }

        private Boolean ImportTWBfromXmlDocument(XmlDocument myXPathDoc)
        {

            String xslPath = "";
            if (Properties.Settings.Default.ImportConvertersAutoSearch)
            {
                String version = myXPathDoc.SelectSingleNode("/workbook/@version").InnerText;
                String converter = Properties.Settings.Default.ImportConvertersLocalization + "\\TRDCI_v" + version + ".xsl";
                if (File.Exists(converter))
                {
                    xslPath = converter;
                }
                if ("".Equals(xslPath) && File.Exists(Properties.Settings.Default.ImportConverterDefaultInstance))
                {
                    xslPath = Properties.Settings.Default.ImportConverterDefaultInstance;
                }
            }
            else
            {
                if (File.Exists(Properties.Settings.Default.UserConverter))
                {
                    xslPath = Properties.Settings.Default.UserConverter;
                }
            }

            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            if (!"".Equals(xslPath) && File.Exists(xslPath))
            {
                myXslTrans.Load(xslPath);
            }
            else
            {
                throw new Exception("\".twb\" File Converter in Path:"+xslPath+" Not Found!");
            }

            converted = new XmlDocument();
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
                                var doc = new XmlDocument();
                                doc.Load(filename);
                                return ImportTWBfromXmlDocument(doc);
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
