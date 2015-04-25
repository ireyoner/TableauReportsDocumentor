﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using TableauReportsDocumentor.Modules.ImportModule;

namespace TableauReportsDocumentor.Data
{
    class ReportDocument
    {
        private XmlDocument xml;
        private String fileName;
        private String directoryName;
        private ImportTWBandTWBX importTWBandTWBX;
        private ValidationEventHandler veh;

        public String FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                if (value != null)
                {
                    this.fileName = Path.GetFileName(value);
                }
                else
                {
                    this.fileName = null;
                }
            }
        }
        public String DirectoryName
        {
            get
            {
                return directoryName;
            }
            set
            {
                if (value != null)
                {
                    this.directoryName = Path.GetDirectoryName(value);
                }
                else
                {
                    this.directoryName = null;
                }
            }
        }
        public String FullFilePath
        {
            get
            {
                if (fileName != null)
                    return directoryName +'\\'+ fileName;
                else
                    return null;
            }
            set
            {
                if (value != null)
                {
                    this.fileName = Path.GetFileName(value);
                    this.directoryName = Path.GetDirectoryName(value);
                }
                else
                {
                    this.fileName = null;
                    this.directoryName = null;
                }
            }
        }
        
        public XmlDocument Xml
        {
            get
            {
                return xml;
            }
            set 
            {
                if (value != null)
                {
                    //var old_xml = xml;
                    xml = value;
                    try
                    {
                        ValideteXML();
                    }catch(Exception e){
                        //xml = old_xml;
                        throw e;
                    }
                }
            }
        }

        public ReportDocument()
        {
            SetupReportDocument();
        }

        public ReportDocument(XmlDocument xmlDocument)
        {
            SetupReportDocument();
            this.Xml = xmlDocument;
        }

        private void SetupReportDocument()
        {
            DirectoryName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            DirectoryName = @"C:\Projects\trd temp files\";
            this.xml = new XmlDocument();
            veh = new ValidationEventHandler(ValidationCallBack);
            this.importTWBandTWBX = new ImportTWBandTWBX();
        }

        public bool Open()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook, Tableau Report Documentator|*.trd;*.twb;*.twbx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = this.DirectoryName;

            if (openFileDialog.ShowDialog() == true)
            {
                this.DirectoryName = openFileDialog.FileName;

                string filename = openFileDialog.FileName;

                if (filename.EndsWith(".twb") || filename.EndsWith(".twbx"))
                {
                    this.Xml = importTWBandTWBX.ImportTableauWorkbook(filename);
                }
                else
                {
                    if (filename.EndsWith(".trd"))
                    {
                        this.FileName = openFileDialog.FileName;
                    }
                    var newXml = new XmlDocument();
                    newXml.Load(filename);
                    this.Xml = newXml;
                }
                return true;
            }
            return false;
        }

        public bool SaveAs()
        {
            return Save(true);
        }

        public bool Save()
        {
            return Save(false);
        }

        private bool Save(bool asNewInstance)
        {
            if (this.FileName == null || asNewInstance)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Tableau Report Documentator (*.trd)|*.trd|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "trd";
                saveFileDialog.InitialDirectory = directoryName;
                if (saveFileDialog.ShowDialog() ?? false)
                {
                    this.FullFilePath = saveFileDialog.FileName;
                    Console.WriteLine(this.FullFilePath);
                    this.xml.Save(this.FullFilePath);
                    return true;
                }
            }
            else
            {
                this.xml.Save(this.FullFilePath);
                return true;
            }
            return false;
        }

        public String GetAsString()
        {
            if (Xml != null)
            {
                MemoryStream mStream = new MemoryStream();
                XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
                writer.Formatting = Formatting.Indented;
                writer.Indentation = 4;
                writer.QuoteChar = '\'';
                this.xml.WriteContentTo(writer);
                writer.Flush();
                mStream.Flush();

                // Have to rewind the MemoryStream in order to read
                // its contents.
                mStream.Position = 0;

                // Read MemoryStream contents into a StreamReader.
                StreamReader sReader = new StreamReader(mStream);

                // Extract the text from the StreamReader.
                String FormattedXML = sReader.ReadToEnd();

                return FormattedXML;
            }else{
                return "";
            }
        }

        public bool SaveFromString(String stringXml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(stringXml);
            this.Xml = doc;
            return true;
        }

        public bool ValideteXML()
        {
            if (Xml.Schemas.Count < 1)
            {
                Xml.Schemas.Add("", "../../Import Converters/ImportValidator.xsd");
            }
            Xml.Validate(veh);
            return true;
        }

        private static void ValidationCallBack(object sender, ValidationEventArgs args)
        {
            String message = "";
            if (args.Severity == XmlSeverityType.Warning)
            {
                message = "\tWarning: Matching schema not found.  No validation occurred." + args.Message;
            }
            else
            {
                message = "\tValidation error: " + args.Message;
            }
            Console.WriteLine(message);
            throw new Exception("Document validation exception!\n" + message);
        }


        public XmlDocument GetExportXml()
        {
            Console.WriteLine("e1:");
            XmlDocument doc = (XmlDocument)this.Xml.Clone();
            var elem = doc.SelectNodes("//node()[@visible='False']");
            while ( elem.Count > 0)
            {
                Console.WriteLine("e:"+elem[0].Name+" ");
                elem[0].ParentNode.RemoveChild(elem[0]);
                elem = doc.SelectNodes("//node()[@visible='False']");
            }
            return doc;
        }
    }
}
