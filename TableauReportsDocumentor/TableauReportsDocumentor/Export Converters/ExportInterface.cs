using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;

namespace TableauReportsDocumentor.Export_Converters
{
    interface ExportInterface
    {
        String displayName { get; }
        String exportFormat { get; }
        String exportFormatFilter { get; }
        BitmapFrame icone { get; }

        bool export(String exportFileName, XmlDocument exportSource);
    }
}
