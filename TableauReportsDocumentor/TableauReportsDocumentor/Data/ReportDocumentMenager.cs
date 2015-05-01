using Microsoft.Win32;
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
    public class ReportDocumentMenager
    {
        private ReportContent content;
        public ReportContent Content
        {
            get { return content; }
            set { content = value; }
        }

        private String fileName;
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
        
        private String directoryName;
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
                    return directoryName + '\\' + fileName;
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

        private ImportTWBandTWBX importTWBandTWBX;

        public ReportDocumentMenager()
        {
            SetupReportDocument(new ReportContent());
        }

        public ReportDocumentMenager(XmlDocument xmlDocument)
        {
            SetupReportDocument(new ReportContent(xmlDocument));
        }

        public ReportDocumentMenager(ReportContent reportContent)
        {
            SetupReportDocument(reportContent);
        }

        private void SetupReportDocument(ReportContent reportContent)
        {
            DirectoryName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            DirectoryName = @"C:\Projects\trd temp files\";
            this.Content = reportContent;
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
                this.FullFilePath = openFileDialog.FileName;
                if (FullFilePath.EndsWith(".twb") || FullFilePath.EndsWith(".twbx"))
                {
                    Boolean status = importTWBandTWBX.ImportTableauWorkbook(FullFilePath);         
                    if (status)
                    {
                        this.Content = new ReportContent(importTWBandTWBX.OriginalReport, importTWBandTWBX.ConvertedReport);
                    }
                    return status;
                }
                else
                {
                    if (FullFilePath.EndsWith(".trd"))
                    {
                        this.FileName = openFileDialog.FileName;
                    }
                    var newXml = new XmlDocument();
                    newXml.Load(FullFilePath);
                    this.Content = new ReportContent(newXml);
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
                    this.Content.ConvertedXml.Save(this.FullFilePath);
                    return true;
                }
            }
            else
            {
                this.Content.ConvertedXml.Save(this.FullFilePath);
                return true;
            }
            return false;
        }

        public XmlDocument GetExportXml()
        {
            return this.Content.GetExportXml();
        }
    }
}
