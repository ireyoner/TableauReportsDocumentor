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
        public String name { get; set; }
        public String path { get; set; }

        public ReportDocument()
        {
            path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            this.xml = new XmlDocument();
        }

        public bool Open()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Tableau Workbook, Tableau Report Documentator|*.trd;*.twb;*.twbx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = path;

            if (openFileDialog.ShowDialog() == true)
            {
                this.name = openFileDialog.SafeFileName;
                this.path = openFileDialog.FileName;

                string filename = openFileDialog.FileName;

                if (filename.EndsWith(".trd"))
                {
                    this.xml.Load(filename);
                }
                else
                {
                    var imp = new Import();
                    this.xml = imp.ImportTableauWorkbooks(filename);
                }
                return true;
            }
            return false;
        }

        public bool Save()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Tableau Report Documentator (*.trd)|*.trd|All files (*.*)|*.*";
            saveFileDialog.InitialDirectory = path;

            if (saveFileDialog.ShowDialog() == true)
            {
                this.name = saveFileDialog.SafeFileName;
                this.path = saveFileDialog.FileName;

                string filename = saveFileDialog.FileName;

                this.xml.Save(filename);
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

        public bool saveFromString(String stringXml){
            return false;
        }

    }
}
