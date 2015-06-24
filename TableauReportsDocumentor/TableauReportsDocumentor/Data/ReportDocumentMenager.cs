using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using TableauReportsDocumentor.Modules.ImportModule;

namespace TableauReportsDocumentor.Data
{
    public class ReportDocumentManager
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
                if (directoryName != null && fileName != null)
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

        public ReportDocumentManager()
        {
            SetupReportDocument(new ReportContent());
        }

        public ReportDocumentManager(XmlDocument xmlDocument)
        {
            SetupReportDocument(new ReportContent(xmlDocument));
        }

        public ReportDocumentManager(ReportContent reportContent)
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
            openFileDialog.Filter = "Tableau Workbook, Tableau Report Documentator|*.trdx;*.twb;*.twbx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = this.DirectoryName;

            if (openFileDialog.ShowDialog() == true)
            {
                this.FullFilePath = null;
                if (openFileDialog.FileName.EndsWith(".twb") || openFileDialog.FileName.EndsWith(".twbx"))
                {
                    Boolean status = importTWBandTWBX.ImportTableauWorkbook(openFileDialog.FileName);
                    if (status)
                    {
                        this.Content = new ReportContent(importTWBandTWBX.OriginalReport, importTWBandTWBX.ConvertedReport);
                        this.FileName = openFileDialog.FileName;
                    }
                    return status;
                }
                else if (openFileDialog.FileName.EndsWith(".trdx"))
                {
                    this.FullFilePath = openFileDialog.FileName;
                    var trdxXml = new XmlDocument();
                    String originalTwb = "";

                    FileStream appArchiveFileStream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                    using (ZipArchive archive = new ZipArchive(appArchiveFileStream, ZipArchiveMode.Read))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith("original.twb", StringComparison.OrdinalIgnoreCase))
                            {
                                using (Stream appManifestFileStream = entry.Open())
                                {
                                    StreamReader streamReader = new StreamReader(appManifestFileStream);
                                    originalTwb = streamReader.ReadToEnd();
                                    streamReader.Close();
                                }
                            }
                            else if (entry.FullName.EndsWith("report.trd", StringComparison.OrdinalIgnoreCase))
                            {
                                using (Stream appManifestFileStream = entry.Open())
                                {
                                    StreamReader streamReader = new StreamReader(appManifestFileStream);
                                    trdxXml.Load(streamReader);
                                    streamReader.Close();
                                }
                            }
                        }
                    }

                    this.Content = new ReportContent(originalTwb, trdxXml);
                    return true;
                }
                else
                {
                    var newXml = new XmlDocument();
                    newXml.Load(openFileDialog.FileName);
                    this.Content = new ReportContent(newXml);
                    this.FileName = openFileDialog.FileName;
                    return true;
                }
            }
            else
            {
                return false;
            }
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
            if (this.FullFilePath == null || asNewInstance)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Tableau Report Documentator (*.trdx)|*.trdx|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "trdx";
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(FileName);
                if (saveFileDialog.ShowDialog() ?? false)
                {
                    return SaveTRDX(saveFileDialog.FileName);
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return SaveTRDX(this.FullFilePath);
            }
        }

        private bool SaveTRDX(String filepath)
        {
            File.Delete(filepath);
            using (FileStream zipToOpen = new FileStream(filepath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    ZipArchiveEntry originalTwb = archive.CreateEntry("original.twb");
                    using (StreamWriter writer = new StreamWriter(originalTwb.Open()))
                    {
                        writer.Write(this.Content.Original);
                    }
                    ZipArchiveEntry trdXml = archive.CreateEntry("report.trd");
                    using (StreamWriter writer = new StreamWriter(trdXml.Open()))
                    {
                        writer.Write(this.Content.ConvertedXmlAsString);
                    }
                }
            }
            Console.WriteLine(filepath);
            this.FullFilePath = filepath;
            return true;
        }

        public XmlDocument GetExportXml()
        {
            return this.Content.GetExportXml();
        }
    }
}
