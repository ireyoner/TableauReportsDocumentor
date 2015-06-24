/*
 * File: ImportTWBandTWBX.cs
 * Class: ImportTWBandTWBX
 * 
 * Class responsible for importing and converting data from TWB TWBX files to Report XmlDocument
 * 
 */
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

        /*
         * Function: ImportTWB
         * Import data from TWB file
         * 
         * Parameters:
         *  filename - a path to file to import
         *  
         * Returns:
         *  if a import was succesfull 
         */
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

        /*
         * Function: ImportTWBfromXmlDocument
         * Function responsible for converting original Tableau xml to report xml using xls converter.
         * 
         * Parameters:
         *  myXPathDoc - original Tableau xml
         *  
         * Returns:
         *  if a conversion was succesfull 
         */
        private Boolean ImportTWBfromXmlDocument(XmlDocument myXPathDoc)
        {

            String xslPath = "";
            // when auto search
            if (Properties.Settings.Default.ImportConvertersAutoSearch)
            {
                //find original report version
                String version = myXPathDoc.SelectSingleNode("/workbook/@version").InnerText;

                // find valid converter for version
                String converter = Properties.Settings.Default.ImportConvertersLocalization + "\\TRDCI_v" + version + ".xsl";
                if (File.Exists(converter))
                {
                    xslPath = converter;
                }
                // if not exist valid converter for fersion get fallback converter
                if ("".Equals(xslPath) && File.Exists(Properties.Settings.Default.ImportConverterDefaultInstance))
                {
                    xslPath = Properties.Settings.Default.ImportConverterDefaultInstance;
                }
            }
            // when user defined converter
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

        /*
         * Function: ImportTWBX
         * Import data from TWBX file
         * 
         * Parameters:
         *  filename - a path to file to import
         *  
         * Returns:
         *  if a import was succesfull 
         */
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

        /*
         * Function: filename
         * calls specific importer for file version
         * 
         * Parameters:
         *  filename - a path to file to import
         *  
         * Returns:
         *  if a import was succesfull 
         */
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
