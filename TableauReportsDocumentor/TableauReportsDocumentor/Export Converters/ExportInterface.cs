using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TableauReportsDocumentor.Export_Converters
{
    interface ExportInterface
    {
        String exportFormat { get; }
        bool export(String exportFileName, XmlDocument exportSource);
    }
}
