using Microsoft.Win32;
using System;
using System.Collections.Generic;
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


    }
}
