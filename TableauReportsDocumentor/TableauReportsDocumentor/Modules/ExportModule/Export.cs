using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using TableauReportsDocumentor.Export_Converters;

namespace TableauReportsDocumentor.Modules.ExportModule
{
    class Export
    {

        public bool ExportTableauReportDocumentatnion(XmlDocument outputDoc){

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Docx document|*.docx";
            saveFileDialog.Title = "Save prepared aport documentation";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != "")
            {
                ExportInterface exporter;
                exporter = new DocxExport();
                return exporter.export(saveFileDialog.FileName, outputDoc);

            }
            return false;
        }
    }
}
