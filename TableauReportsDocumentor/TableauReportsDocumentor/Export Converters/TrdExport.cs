using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableauReportsDocumentor.Export_Converters
{
    class TrdExport : ExportInterface
    {
        public string MenuItemText
        {
            get { return null; }
        }

        public string FileExtinsion
        {
            get { return "trd"; }
        }

        public string ExportSaveFileDialogFilter
        {
            get { return null; }
        }

        public System.Windows.Media.Imaging.BitmapFrame MenuItemIcone
        {
            get { return null; }
        }

        public object ToolBarButtonContent
        {
            get { return null; }
        }

        public bool Export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            exportSource.Save(exportFileName);
            return true;
        }
    }
}
