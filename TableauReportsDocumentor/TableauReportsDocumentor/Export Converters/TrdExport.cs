using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableauReportsDocumentor.Export_Converters
{
    class TrdExport : ExportInterface
    {
        public string menuItemText
        {
            get { return null; }
        }

        public string fileExtinsion
        {
            get { return "trd"; }
        }

        public string exportSaveFileDialogFilter
        {
            get { return null; }
        }

        public System.Windows.Media.Imaging.BitmapFrame menuItemIcone
        {
            get { return null; }
        }

        public object toolBarButtonContent
        {
            get { return null; }
        }

        public bool export(string exportFileName, System.Xml.XmlDocument exportSource)
        {
            exportSource.Save(exportFileName);
            return true;
        }
    }
}
