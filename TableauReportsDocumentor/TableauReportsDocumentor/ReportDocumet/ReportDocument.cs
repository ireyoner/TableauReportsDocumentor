using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using TableauReportsDocumentor.Modules.ImportModule;

namespace TableauReportsDocumentor.ReportDocumet
{
    class ReportDocument
    {
        public XmlDocument xml { get; private set; }
        private String fileName;
        private String directoryName;
        private Import import;

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
                    return directoryName + fileName;
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

        public ReportDocument()
        {
            DirectoryName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            this.xml = new XmlDocument();
            this.import = new Import();
        }

        public bool Open()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook, Tableau Report Documentator|*.trd;*.twb;*.twbx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = DirectoryName;

            if (openFileDialog.ShowDialog() == true)
            {
                this.DirectoryName = Path.GetDirectoryName(openFileDialog.FileName);

                string filename = openFileDialog.FileName;

                if (filename.EndsWith(".twb") || filename.EndsWith(".twbx"))
                {
                    this.xml = import.ImportTableauWorkbooks(filename);
                }
                else
                {
                    if (filename.EndsWith(".trd"))
                    {
                        this.FileName = openFileDialog.FileName;
                    }
                    this.xml.Load(filename);
                }
                return true;
            }
            return false;
        }

        public bool SaveAs()
        {
            return Save(false);
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
                if (saveFileDialog.ShowDialog() == true)
                {
                    this.FullFilePath = saveFileDialog.FileName;
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

        public String getAsString()
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
        }

        public bool saveFromString(String stringXml)
        {
            return false;
        }

    }
}
