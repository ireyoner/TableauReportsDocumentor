using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TableauReportsDocumentor.Data
{
    public class ImportedDocument {
    public ImportedDocument() {
    }

    public ImportedDocument(XmlDocument original, XmlDocument converted)
    {
        this.Original = original;
        this.Converted = converted;
    }

    public XmlDocument Original { get; set; }
    public XmlDocument Converted { get; set; }
};
}
